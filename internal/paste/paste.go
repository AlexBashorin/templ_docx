package paste

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"image/png"
	"io"
	"io/ioutil"
	"log"
	"regexp"
	"strings"
)

type Document struct {
	XMLName xml.Name `xml:"document"`
	Body    Body     `xml:"body"`
}

type Body struct {
	Paragraphs []Paragraph `xml:"p"`
}

type Paragraph struct {
	Content []interface{} `xml:",any"`
}

type Run struct {
	XMLName       xml.Name      `xml:"r"`
	RunProperties RunProperties `xml:"rPr"`
	Text          Text          `xml:"t"`
}

type Hyperlink struct {
	XMLName xml.Name `xml:"hyperlink"`
	ID      string   `xml:"id,attr"`
	Run     Run      `xml:"r"`
}

type RunProperties struct {
	Fonts  Fonts  `xml:"rFonts"`
	Size   Size   `xml:"sz"`
	SizeCs SizeCs `xml:"szCs"`
	Color  Color  `xml:"color,omitempty"`
}

type Fonts struct {
	ASCII string `xml:"ascii,attr"`
	HANSI string `xml:"hAnsi,attr"`
}

type Size struct {
	Val string `xml:"val,attr"`
}

type SizeCs struct {
	Val string `xml:"val,attr"`
}

type Text struct {
	Value string `xml:",chardata"`
}

type Color struct {
	Val string `xml:"val,attr"`
}

func ReplaceVariablesInDOCX(docxFile []byte, variables map[string]string, images map[string][]byte, links, qrcodes map[string]string) ([]byte, error) {
	reader, err := zip.NewReader(bytes.NewReader(docxFile), int64(len(docxFile)))
	if err != nil {
		return nil, fmt.Errorf("error opening docx: %v", err)
	}

	var resultBuf bytes.Buffer
	writer := zip.NewWriter(&resultBuf)

	var documentXML []byte
	var relsXML []byte
	imageRelationships := make(map[string]string)
	linksRelationships := make(map[string]string)
	QRRelationships := make(map[string]string)

	// Convert xml to structure
	var xmldoc Document

	for _, file := range reader.File {
		switch file.Name {
		case "word/document.xml":
			documentXML, err = readZipFile(file)
			if err != nil {
				return nil, fmt.Errorf("error reading document.xml: %v", err)
			}
			errXmlUnmarsh := xml.Unmarshal(documentXML, &xmldoc)
			if errXmlUnmarsh != nil {
				return nil, fmt.Errorf("Can't unmarshal xml")
			}
			log.Printf("document.xml content length: %d", len(documentXML))
		case "word/_rels/document.xml.rels":
			relsXML, err = readZipFile(file)
			if err != nil {
				return nil, fmt.Errorf("error reading document.xml.rels: %v", err)
			}
			log.Printf("document.xml.rels content length: %d", len(relsXML))
		default:
			if err = copyZipFile(writer, file); err != nil {
				return nil, fmt.Errorf("error copying file %s: %v", file.Name, err)
			}
		}
	}

	// images
	log.Printf("Processing %d images", len(images))
	for key, imageData := range images {
		rId := fmt.Sprintf("rId%d", len(imageRelationships)+1000)
		fileName := fmt.Sprintf("image%d.png", len(imageRelationships)+1)
		imageRelationships[key] = rId

		log.Printf("Adding image: %s with rId: %s", fileName, rId)

		imgWriter, err := writer.Create(fmt.Sprintf("word/media/%s", fileName))
		if err != nil {
			return nil, fmt.Errorf("error creating image file %s: %v", fileName, err)
		}
		if _, err = imgWriter.Write(imageData); err != nil {
			return nil, fmt.Errorf("error writing image data for %s: %v", fileName, err)
		}

		newRel := createImageRelationship(rId, fileName)
		relsXML = []byte(strings.Replace(string(relsXML), "</Relationships>", newRel+"</Relationships>", 1))
		log.Printf("Added relationship: %s", newRel)
	}

	log.Printf("Replacing image variables")
	for key, rId := range imageRelationships {
		pattern := regexp.MustCompile(`<w:t>` + regexp.QuoteMeta(key) + "</w:t>")
		replacement := generateImageXML(rId)
		documentXML = pattern.ReplaceAll(documentXML, replacement)
		log.Printf("Replaced @%s with image XML (rId: %s)", key, rId)
	}

	// texts
	log.Printf("Replacing text variables")
	for key, value := range variables {
		pattern := regexp.MustCompile(regexp.QuoteMeta(key))
		documentXML = pattern.ReplaceAll(documentXML, []byte(value))
		log.Printf("Replaced @%s with %s", key, value)
	}

	// links
	for key, linkURL := range links {
		rId := fmt.Sprintf("rId%d", len(linksRelationships)+1001)
		linksRelationships[key] = rId

		newRel := createLinkRelationship(rId, linkURL)
		relsXML = []byte(strings.Replace(string(relsXML), "</Relationships>", newRel+"</Relationships>", 1))
	}

	for i, para := range xmldoc.Body.Paragraphs {
		for j, content := range para.Content {
			if run, ok := content.(Run); ok {
				if run.Text.Value == "link3d" {
					hyperlink := Hyperlink{
						ID: "rId1", // Это ID должно соответствовать ID в файле word/_rels/document.xml.rels
						Run: Run{
							RunProperties: RunProperties{
								Fonts:  run.RunProperties.Fonts,
								Size:   run.RunProperties.Size,
								SizeCs: run.RunProperties.SizeCs,
								Color:  Color{Val: "0000FF"}, // Синий цвет для ссылки
							},
							Text: Text{Value: "link_to_3d"},
						},
					}
					xmldoc.Body.Paragraphs[i].Content[j] = hyperlink
				}
			}
		}
	}
	output, err := xml.MarshalIndent(xmldoc, "", "  ")
	if err != nil {
		panic(err)
	}

	// Замена пространства имен
	output = bytes.Replace(output, []byte("document"), []byte("w:document"), 1)
	output = bytes.Replace(output, []byte("<body>"), []byte("<w:body>"), 1)
	output = bytes.Replace(output, []byte("</body>"), []byte("</w:body>"), 1)
	output = bytes.Replace(output, []byte("<p>"), []byte("<w:p>"), -1)
	output = bytes.Replace(output, []byte("</p>"), []byte("</w:p>"), -1)
	output = bytes.Replace(output, []byte("<r>"), []byte("<w:r>"), -1)
	output = bytes.Replace(output, []byte("</r>"), []byte("</w:r>"), -1)
	output = bytes.Replace(output, []byte("<rPr>"), []byte("<w:rPr>"), -1)
	output = bytes.Replace(output, []byte("</rPr>"), []byte("</w:rPr>"), -1)
	output = bytes.Replace(output, []byte("<rFonts"), []byte("<w:rFonts"), -1)
	output = bytes.Replace(output, []byte("<sz"), []byte("<w:sz"), -1)
	output = bytes.Replace(output, []byte("<szCs"), []byte("<w:szCs"), -1)
	output = bytes.Replace(output, []byte("<t>"), []byte("<w:t>"), -1)
	output = bytes.Replace(output, []byte("</t>"), []byte("</w:t>"), -1)
	output = bytes.Replace(output, []byte("<hyperlink"), []byte("<w:hyperlink"), -1)
	output = bytes.Replace(output, []byte("</hyperlink>"), []byte("</w:hyperlink>"), -1)
	output = bytes.Replace(output, []byte("<color"), []byte("<w:color"), -1)

	// Добавление пространства имен
	output = append([]byte(xml.Header), output...)
	output = append([]byte(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+"\n"), output...)
	output = append(output, []byte("\n</w:document>")...)

	// QR code
	for key, qrdata := range qrcodes {
		// set to rels
		rIdQR := fmt.Sprintf("rId%d", len(QRRelationships)+9000)
		relQRxml := createQRRelationship(rIdQR, "qr.png")
		relsXML = []byte(strings.Replace(string(relsXML), "</Relationships>", relQRxml+"</Relationships>", 1))
		QRRelationships[key] = rIdQR

		genQRCode := GenerateQRCode(qrdata, 256)
		imgWriter, err := writer.Create(fmt.Sprintf("word/media/%s", key))
		if err != nil {
			return nil, fmt.Errorf("error creating image file %s: %v", key, err)
		}
		buf := new(bytes.Buffer)
		errBytesImage := png.Encode(buf, genQRCode)
		if errBytesImage != nil {
			log.Fatalf("errBytesImage: %v\n", errBytesImage)
		}
		if _, err = imgWriter.Write(buf.Bytes()); err != nil {
			return nil, fmt.Errorf("error writing image data for %s: %v", key, err)
		}
	}

	// QR rels
	for key, rId := range linksRelationships {
		pattern := regexp.MustCompile("<w:r><w:rPr></w:rPr><w:t>" + regexp.QuoteMeta(key) + "</w:t></w:r>")
		replacement := generateLinkXML(rId, key)
		fmt.Printf("generateLinkXML: %v\n", replacement)
		documentXML = pattern.ReplaceAll(documentXML, replacement)
	}

	// var qrBuffer bytes.Buffer
	// png.Encode(&qrBuffer, genQRCode)
	// qrBase64 := base64.StdEncoding.EncodeToString(qrBuffer.Bytes())

	// pattern := regexp.MustCompile(`<w:t>` + regexp.QuoteMeta(key) + "</w:t>")
	// replacement := generateImageXML(rIdQR)
	// documentXML = pattern.ReplaceAll(documentXML, replacement)

	// writes
	log.Printf("Writing updated document.xml")
	if err = writeZipFile(writer, "word/document.xml", documentXML); err != nil {
		return nil, fmt.Errorf("error writing document.xml: %v", err)
	}

	log.Printf("Writing updated document.xml.rels")
	if err = writeZipFile(writer, "word/_rels/document.xml.rels", relsXML); err != nil {
		return nil, fmt.Errorf("error writing document.xml.rels: %v", err)
	}

	err = writer.Close()
	if err != nil {
		return nil, fmt.Errorf("error closing zip writer: %v", err)
	}

	log.Printf("DOCX processing completed")
	return resultBuf.Bytes(), nil
}

func generateImageXML(rId string) []byte {
	return []byte(fmt.Sprintf(`
	<w:drawing>
		<wp:anchor behindDoc="0" distT="0" distB="0" distL="0" distR="0" simplePos="0"
			locked="0" layoutInCell="0" allowOverlap="1" relativeHeight="2">
			<wp:simplePos x="0" y="0" />
			<wp:positionH relativeFrom="column">
				<wp:posOffset>-24130</wp:posOffset>
			</wp:positionH>
			<wp:positionV relativeFrom="paragraph">
				<wp:posOffset>88900</wp:posOffset>
			</wp:positionV>
			<wp:extent cx="2592070" cy="2592070" />
			<wp:effectExtent l="0" t="0" r="0" b="0" />
			<wp:wrapSquare wrapText="largest" />
			<wp:docPr id="1" name="Image1" descr=""></wp:docPr>
			<wp:cNvGraphicFramePr>
				<a:graphicFrameLocks
					xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
					noChangeAspect="1" />
			</wp:cNvGraphicFramePr>
			<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
				<a:graphicData
					uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:pic
						xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:nvPicPr>
							<pic:cNvPr id="1" name="Image1" descr=""></pic:cNvPr>
							<pic:cNvPicPr>
								<a:picLocks noChangeAspect="1" noChangeArrowheads="1" />
							</pic:cNvPicPr>
						</pic:nvPicPr>
						<pic:blipFill>
							<a:blip r:embed="%s"></a:blip>
							<a:stretch>
								<a:fillRect />
							</a:stretch>
						</pic:blipFill>
						<pic:spPr bwMode="auto">
							<a:xfrm>
								<a:off x="0" y="0" />
								<a:ext cx="2592070" cy="2592070" />
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst />
							</a:prstGeom>
						</pic:spPr>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		</wp:anchor>
	</w:drawing>`, rId))
}

func generateQRXML(rId string) []byte {
	return []byte(fmt.Sprintf(`
		<w:drawing>
			<wp:inline>
				<wp:extent cx="2700000" cy="2700000"/>
				<wp:docPr id="1" name="QR Code"/>
				<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
					<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
							<pic:blipFill>
								<a:blip r:embed="%s"/>
							</pic:blipFill>
						</pic:pic>
					</a:graphicData>
				</a:graphic>
			</wp:inline>
		</w:drawing>
`, rId))
}

func generateLinkXML(rId, textlink string) []byte {
	return []byte(fmt.Sprintf(`<w:hyperlink r:id="%s">
		<w:r>
			<w:rPr>
				<w:rStyle w:val="Hyperlink" />
			</w:rPr>
			<w:t>%s</w:t>
		</w:r>
	</w:hyperlink>`, rId, textlink))
}

func createImageRelationship(rId, fileName string) string {
	return fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, rId, fileName)
}

func createLinkRelationship(rId, linkURL string) string {
	return fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External" />`, rId, linkURL)
}

func createQRRelationship(rId, qrname string) string {
	return fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, rId, qrname)
}

func readZipFile(file *zip.File) ([]byte, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	return ioutil.ReadAll(rc)
}

func writeZipFile(writer *zip.Writer, filename string, content []byte) error {
	w, err := writer.Create(filename)
	if err != nil {
		return err
	}

	_, err = w.Write(content)
	return err
}

func copyZipFile(writer *zip.Writer, file *zip.File) error {
	rc, err := file.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	w, err := writer.Create(file.Name)
	if err != nil {
		return err
	}

	_, err = io.Copy(w, rc)
	return err
}
