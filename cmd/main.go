package main

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"image/png"
	"io"
	"log"
	"net/http"
	"regexp"
	"strings"

	"github.com/boombuler/barcode"
	"github.com/boombuler/barcode/qr"
)

// DocxTemplate представляет структуру DOCX файла
type DocxTemplate struct {
	files map[string][]byte
}

// TemplateData содержит данные для замены в шаблоне
type TemplateData struct {
	Text    map[string]string     `json:"text"`
	Images  map[string]ImageData  `json:"images"`
	QRCodes map[string]string     `json:"qrcodes"`
	Tables  map[string][][]string `json:"tables"`
	Links   map[string]string     `json:"links"`
}

// ImageData содержит информацию об изображении
type ImageData struct {
	Data   string `json:"data"` // Base64-encoded image data
	Width  int    `json:"width"`
	Height int    `json:"height"`
}

type WordML struct {
	XMLName xml.Name `xml:"document"`
	Body    Body     `xml:"body"`
}

type Body struct {
	XMLName    xml.Name    `xml:"body"`
	Paragraphs []Paragraph `xml:"w:p"`
}

type Paragraph struct {
	XMLName xml.Name `xml:"w:p"`
	Run     []Run    `xml:"w:r"`
}

type Run struct {
	XMLName       xml.Name      `xml:"w:r"`
	RunProperties RunProperties `xml:"w:rPr"`
	Text          string        `xml:"w:t"`
	Hyperlink     Hyperlink     `xml:"w:hyperlink"`
}

type RunProperties struct {
	XMLName xml.Name `xml:"w:rPr"`
}

type Hyperlink struct {
	XMLName xml.Name `xml:"w:hyperlink"`
	RID     string   `xml:"r:id,attr"`
}

// LoadTemplate загружает DOCX файл как шаблон
func LoadTemplate(data []byte) (*DocxTemplate, error) {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}

	template := &DocxTemplate{
		files: make(map[string][]byte),
	}

	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, err
		}
		template.files[file.Name] = content
	}

	return template, nil
}

// ApplyTemplate применяет данные к шаблону
func (t *DocxTemplate) ApplyTemplate(data TemplateData) error {
	// Обработка document.xml
	docXml, err := t.getFileContent("word/document.xml")
	if err != nil {
		return err
	}

	// Замена текста
	// log.Printf("Text")
	if data.Text != nil {
		for key, value := range data.Text {
			docXml = strings.ReplaceAll(docXml, key, value)
		}
	} else {
		log.Printf("NIL Text")
	}

	// Замена изображений
	// log.Printf("Images")
	if data.Images != nil {
		ind := 1000
		for key, imgData := range data.Images {
			ind += 1
			rId := fmt.Sprintf("rId%d", ind)
			fileName := fmt.Sprintf("image%d.png", ind)
			// imageRelationships[key] = rId

			// Создание зависимостей
			newRel := createImageRelationship(rId, fileName)
			relsContent, err := t.getFileContent("word/_rels/document.xml.rels")
			if err != nil {
				return err
			}
			relsContent = strings.Replace(relsContent, "</Relationships>", newRel+"</Relationships>", 1)
			t.files["word/_rels/document.xml.rels"] = []byte(relsContent)
			// relsXML = []byte(strings.Replace(string(relsXML), "</Relationships>", newRel+"</Relationships>", 1))

			// placeholder := `<w:rPr><w:rFonts w:ascii="TT Hoves" w:hAnsi="TT Hoves" /><w:sz w:val="22" /><w:szCs w:val="22" /></w:rPr><w:t>` + key + `</w:t>`
			// replacement := ReplaceToImage(x, y, key)
			// replacement := fmt.Sprintf(`<w:drawing><wp:inline><wp:extent cx="%d" cy="%d"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId%s"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`, imgData.Width*9525, imgData.Height*9525, key)

			// Замена XML
			placeholder := `<w:t>` + key + `</w:t>`
			x := imgData.Width * 9525
			y := imgData.Height * 9525
			replacement := generateImageXML(x, y, rId)
			docXml = strings.ReplaceAll(docXml, placeholder, replacement)

			// запись файлов в архив
			// imageId := fmt.Sprintf("rId%s", ind + 1)
			dec, err := base64.StdEncoding.DecodeString(imgData.Data)
			if err != nil {
				panic(err)
			}
			err = t.addImageToArchive(ind, dec)
			if err != nil {
				return err
			}
		}
	} else {
		log.Printf("NIL Images")
	}

	// Замена QR-кодов
	// log.Printf("QR")
	if data.QRCodes != nil {
		qrind := 2000
		for key, value := range data.QRCodes {
			qrCodeBytes, err := generateQRCode(value)
			if err != nil {
				return err
			}

			qrind += 1
			imageId := fmt.Sprintf("rId%d", qrind)
			fileName := fmt.Sprintf("image%d.png", qrind)

			// Создание зависимостей
			newRel := createImageRelationship(imageId, fileName)
			relsContent, err := t.getFileContent("word/_rels/document.xml.rels")
			if err != nil {
				return err
			}
			relsContent = strings.Replace(relsContent, "</Relationships>", newRel+"</Relationships>", 1)
			t.files["word/_rels/document.xml.rels"] = []byte(relsContent)

			// Запись файла
			err = t.addImageToArchive(qrind, qrCodeBytes)
			if err != nil {
				return err
			}

			// Замена XML
			placeholder := `<w:t>` + key + `</w:t>`
			x := 80 * 9525
			y := 80 * 9525
			replacement := generateImageXML(x, y, imageId)
			docXml = strings.ReplaceAll(docXml, placeholder, replacement)
			// replacement := fmt.Sprintf(`<w:drawing><wp:inline><wp:extent cx="%d" cy="%d"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="%s"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`, 30*9525, 30*9525, key)
			// docXml = strings.ReplaceAll(docXml, placeholder, replacement)
		}
	} else {
		log.Printf("NIL QR")
	}

	// Замена таблиц
	// log.Printf("Tables")
	if data.Tables != nil {
		for key, tableData := range data.Tables {
			placeholder := key
			tableXml := generateTableXml(tableData)
			log.Printf("tableXml: %v", tableXml)
			docXml = strings.ReplaceAll(docXml, placeholder, tableXml)
		}
	} else {
		log.Printf("NIL Tables")
	}

	// Links
	// var xmldoc Document
	// errXmlUnmarsh := xml.Unmarshal([]byte(docXml), &xmldoc)
	// if errXmlUnmarsh != nil {
	// 	log.Printf("Can't unmarshal xml")
	// }

	// Разбираем XML
	// var wordML WordML
	// err = xml.Unmarshal([]byte(docXml), &wordML)
	// if err != nil {
	// 	fmt.Println(err)
	// }

	// // Ищем элементы Run с текстом "MY_TEXT"
	// for _, paragraph := range wordML.Body.Paragraphs {
	// 	for _, run := range paragraph.Run {
	// 		if run.Text == "link3d" {
	// 			log.Println("LINK_FOUND!")
	// 			// Создаем гиперссылку
	// 			hyperlink := Hyperlink{
	// 				XMLName: xml.Name{Local: "w:hyperlink"},
	// 				RID:     "rId3001", // ID гиперссылки
	// 			}

	// 			// Заменяем текст на гиперссылку
	// 			run.Text = ""
	// 			run.Hyperlink = hyperlink
	// 		}
	// 	}
	// }

	// // Сериализуем XML обратно в байты
	// marsh, err := xml.Marshal(wordML)
	// if err != nil {
	// 	fmt.Println(err)
	// }
	// docXml = string(marsh)

	if data.Links != nil {
		linkIng := 3000
		for key, value := range data.Links {
			linkIng += 1
			imageId := fmt.Sprintf("rId%d", linkIng)

			// Создание зависимостей
			newRel := createLinkRelationship(imageId, value)
			relsContent, err := t.getFileContent("word/_rels/document.xml.rels")
			if err != nil {
				return err
			}
			relsContent = strings.Replace(relsContent, "</Relationships>", newRel+"</Relationships>", 1)
			t.files["word/_rels/document.xml.rels"] = []byte(relsContent)

			replacement := generateLinkXML(imageId)
			docXml = replaceXML(docXml, key, replacement)

			// Регулярное выражение для поиска <w:r>...</w:r>, содержащего определенное значение в <w:t>
			// re := regexp.MustCompile(`(<w:p>(?:\s*)<w:r>(?:\s*)<w:rPr>(?:.*?)</w:rPr>(?:\s*)<w:t>)` + regexp.QuoteMeta(key) + `(</w:t>(?:\s*)</w:r>(?:\s*)</w:p>)`)
			// replacement := generateLinkXML(imageId)
			// docXml = re.ReplaceAllString(docXml, "${1}"+replacement+"${2}")

			// replacement := generateLinkXML(imageId)
			// placeholder := fmt.Sprintf(`<w:r><w:rPr><w:rFonts w:ascii="TT Hoves" w:hAnsi="TT Hoves" /><w:sz w:val="20" /><w:szCs w:val="20" /></w:rPr><w:t>%s</w:t></w:r>`, key)
			// pattern := regexp.MustCompile(regexp.QuoteMeta(placeholder))
			// docXmlBytes := pattern.ReplaceAll([]byte(docXml), []byte(replacement))

			// pattern := regexp.MustCompile(`<w:r><w:rPr>.*?<w:t>` + key + `</w:t>.*?</w:r>`)
			// replacement := fmt.Sprintf(`<w:r><w:rPr>.*?</w:rPr><w:hyperlink r:id="%s"><w:r><w:rPr>.*?<w:t>%s</w:t>.*?</w:r></w:hyperlink></w:r>`, imageId, "Ссылка на 3D")
			// bdocXml := pattern.ReplaceAll([]byte(docXml), []byte(replacement))
			// docXml = string(bdocXml)

			// reReplace = regexp.MustCompile(reReplace.ReplaceAllString(reReplace.String(), linkReXml))
			// Заменяем текст на гиперссылку
			// bdocx := reReplace.ReplaceAll([]byte(docXml), "")

			// pattern := regexp.MustCompile(`<w:t>` + regexp.QuoteMeta(key) + "</w:t>")
			// replacement := generateImageXML(rId)

			// docXml = string(docXmlBytes)

			// placeholder := `<w:t>` + key + `</w:t>`
			// docXml = strings.ReplaceAll(docXml, placeholder, replacement)
		}
	}
	// for i, para := range xmldoc.Body.Paragraphs {
	// 	for j, content := range para.Content {
	// 		if run, ok := content.(Run); ok {
	// 			if run.Text.Value == "link3d" {
	// 				log.Printf("FOUND_LINK!")
	// 				hyperlink := Hyperlink{
	// 					ID: "rId3001", // Это ID должно соответствовать ID в файле word/_rels/document.xml.rels
	// 					Run: Run{
	// 						RunProperties: RunProperties{
	// 							Fonts:  run.RunProperties.Fonts,
	// 							Size:   run.RunProperties.Size,
	// 							SizeCs: run.RunProperties.SizeCs,
	// 							Color:  Color{Val: "0000FF"}, // Синий цвет для ссылки
	// 						},
	// 						Text: Text{Value: "Ссылка на 3D"},
	// 					},
	// 				}
	// 				xmldoc.Body.Paragraphs[i].Content[j] = hyperlink
	// 			}
	// 		}
	// 	}
	// }
	// output, err := xml.MarshalIndent(xmldoc, "", "  ")
	// if err != nil {
	// 	panic(err)
	// }

	// // Замена пространства имен
	// output = bytes.Replace(output, []byte("document"), []byte("w:document"), 1)
	// output = bytes.Replace(output, []byte("<body>"), []byte("<w:body>"), 1)
	// output = bytes.Replace(output, []byte("</body>"), []byte("</w:body>"), 1)
	// output = bytes.Replace(output, []byte("<p>"), []byte("<w:p>"), -1)
	// output = bytes.Replace(output, []byte("</p>"), []byte("</w:p>"), -1)
	// output = bytes.Replace(output, []byte("<r>"), []byte("<w:r>"), -1)
	// output = bytes.Replace(output, []byte("</r>"), []byte("</w:r>"), -1)
	// output = bytes.Replace(output, []byte("<rPr>"), []byte("<w:rPr>"), -1)
	// output = bytes.Replace(output, []byte("</rPr>"), []byte("</w:rPr>"), -1)
	// output = bytes.Replace(output, []byte("<rFonts"), []byte("<w:rFonts"), -1)
	// output = bytes.Replace(output, []byte("<sz"), []byte("<w:sz"), -1)
	// output = bytes.Replace(output, []byte("<szCs"), []byte("<w:szCs"), -1)
	// output = bytes.Replace(output, []byte("<t>"), []byte("<w:t>"), -1)
	// output = bytes.Replace(output, []byte("</t>"), []byte("</w:t>"), -1)
	// output = bytes.Replace(output, []byte("<hyperlink"), []byte("<w:hyperlink"), -1)
	// output = bytes.Replace(output, []byte("</hyperlink>"), []byte("</w:hyperlink>"), -1)
	// output = bytes.Replace(output, []byte("<color"), []byte("<w:color"), -1)

	// // Добавление пространства имен
	// output = append([]byte(xml.Header), output...)
	// output = append([]byte(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+"\n"), output...)
	// output = append(output, []byte("\n</w:document>")...)

	// Сохранение изменений
	t.files["word/document.xml"] = []byte(docXml)

	return nil
}

func ReplaceToImage(x, y int, rId string) string {
	return fmt.Sprintf(`<w:rPr></w:rPr><w:drawing>
					<wp:anchor behindDoc="0" distT="0" distB="0" distL="0" distR="0"
						simplePos="0" locked="0" layoutInCell="0" allowOverlap="1"
						relativeHeight="3">
						<wp:simplePos x="0" y="0" />
						<wp:positionH relativeFrom="column">
							<wp:align>left</wp:align>
						</wp:positionH>
						<wp:positionV relativeFrom="paragraph">
							<wp:posOffset>635</wp:posOffset>
						</wp:positionV>
						<wp:extent cx="%d" cy="%d" />
						<wp:effectExtent l="0" t="0" r="0" b="0" />
						<wp:wrapSquare wrapText="largest" />
						<wp:docPr id="1" name="Image1" descr=""></wp:docPr>
						<wp:cNvGraphicFramePr>
							<a:graphicFrameLocks
								xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
								noChangeAspect="1" />
						</wp:cNvGraphicFramePr>
						<a:graphic
							xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
							<a:graphicData
								uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
								<pic:pic
									xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
									<pic:nvPicPr>
										<pic:cNvPr id="1" name="Image1" descr=""></pic:cNvPr>
										<pic:cNvPicPr>
											<a:picLocks noChangeAspect="1"
												noChangeArrowheads="1" />
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
											<a:ext cx="%d" cy="%d" />
										</a:xfrm>
										<a:prstGeom prst="rect">
											<a:avLst />
										</a:prstGeom>
									</pic:spPr>
								</pic:pic>
							</a:graphicData>
						</a:graphic>
					</wp:anchor>
				</w:drawing>`, x, y, rId, x, y)
}

func generateImageXML(x, y int, rId string) string {
	return fmt.Sprintf(`
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
			<wp:extent cx="%d" cy="%d" />
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
								<a:ext cx="%d" cy="%d" />
							</a:xfrm>
							<a:prstGeom prst="rect">
								<a:avLst />
							</a:prstGeom>
						</pic:spPr>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		</wp:anchor>
	</w:drawing>`, x, y, rId, x, y)
}

func generateLinkXML(rId string) string {
	return fmt.Sprintf(`<w:hyperlink r:id="%s">
	<w:r>
		<w:rPr>
			<w:rStyle w:val="Hyperlink" />
			<w:rFonts w:ascii="TT Hoves" w:hAnsi="TT Hoves" />
			<w:sz w:val="20" />
			<w:szCs w:val="20" />
		</w:rPr>
		<w:t>Ссылка на 3D</w:t>
	</w:r>
</w:hyperlink>`, rId)
}

func replaceXML(originalXML, keyword, replacementXML string) string {
	// Нормализуем переводы строк в replacementXML
	replacementXML = strings.ReplaceAll(replacementXML, "\n", "")
	replacementXML = strings.ReplaceAll(replacementXML, "\t", "")

	// Компилируем регулярное выражение для поиска целевого <w:r> тега
	re := regexp.MustCompile(`<w:r>.*?<w:t>` + regexp.QuoteMeta(keyword) + `</w:t>.*?</w:r>`)

	// Выполняем замену
	result := re.ReplaceAllStringFunc(originalXML, func(match string) string {
		// Логируем найденное совпадение и замену для отладки
		log.Printf("Found match: %s", match)
		log.Printf("Replacement: %s", replacementXML)

		// Заменяем весь <w:r> тег на новый <w:hyperlink> тег
		return replacementXML
	})

	// Проверяем, была ли выполнена замена
	if result == originalXML {
		log.Println("Warning: No replacements were made. Check your keyword and XML structure.")
	} else {
		log.Println("Replacement successful.")
	}

	return result
}

func (t *DocxTemplate) getFileContent(fileName string) (string, error) {
	content, ok := t.files[fileName]
	if !ok {
		return "", fmt.Errorf("file not found: %s", fileName)
	}
	return string(content), nil
}

func createImageRelationship(rId, fileName string) string {
	return fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, rId, fileName)
}
func (t *DocxTemplate) addImageToArchive(id int, imageData []byte) error {
	imageName := fmt.Sprintf("word/media/image%d.png", id)
	t.files[imageName] = imageData
	// return t.updateRelationships(imageId, imageName)
	return nil
}

func createLinkRelationship(rId, linkURL string) string {
	return fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External" />`, rId, linkURL)
}

// updateRelationships - NOT ACTIVE
func (t *DocxTemplate) updateRelationships(imageId, imageName string) error {
	relsContent, err := t.getFileContent("word/_rels/document.xml.rels")
	if err != nil {
		return err
	}

	newRel := fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="%s"/>`, imageId, imageName)
	relsContent = strings.Replace(relsContent, "</Relationships>", newRel+"</Relationships>", 1)

	t.files["word/_rels/document.xml.rels"] = []byte(relsContent)
	return nil
}

func (t *DocxTemplate) SaveAs() ([]byte, error) {
	buf := new(bytes.Buffer)
	writer := zip.NewWriter(buf)

	for name, content := range t.files {
		f, err := writer.Create(name)
		if err != nil {
			return nil, err
		}
		_, err = f.Write(content)
		if err != nil {
			return nil, err
		}
	}

	err := writer.Close()
	if err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

func generateQRCode(data string) ([]byte, error) {
	qrCode, err := qr.Encode(data, qr.M, qr.Auto)
	if err != nil {
		return nil, err
	}

	qrCode, err = barcode.Scale(qrCode, 200, 200)
	if err != nil {
		return nil, err
	}

	var buf bytes.Buffer
	err = png.Encode(&buf, qrCode)
	if err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

func generateTableXml(data [][]string) string {
	var tableXml strings.Builder
	tableXml.WriteString("<w:tbl>")
	for _, row := range data {
		tableXml.WriteString("<w:tr>")
		for _, cell := range row {
			tableXml.WriteString("<w:tc><w:p><w:r><w:t>")
			tableXml.WriteString(cell)
			tableXml.WriteString("</w:t></w:r></w:p></w:tc>")
		}
		tableXml.WriteString("</w:tr>")
	}
	tableXml.WriteString("</w:tbl>")
	return tableXml.String()
}

func handleTemplate(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Only POST method is allowed", http.StatusMethodNotAllowed)
		return
	}

	w.Write([]byte("GOT"))

	err := r.ParseMultipartForm(10 << 20) // Ограничение в 10 MB
	if err != nil {
		http.Error(w, "Unable to parse form", http.StatusBadRequest)
		return
	}

	file, _, err := r.FormFile("template")
	if err != nil {
		http.Error(w, "Error retrieving the file", http.StatusBadRequest)
		return
	}
	defer file.Close()

	fileBytes, err := io.ReadAll(file)
	if err != nil {
		http.Error(w, "Error reading the file", http.StatusInternalServerError)
		return
	}

	template, err := LoadTemplate(fileBytes)
	if err != nil {
		http.Error(w, "Error creating template", http.StatusInternalServerError)
		return
	}

	var templateData TemplateData
	err = json.Unmarshal([]byte(r.FormValue("data")), &templateData)
	if err != nil {
		http.Error(w, "Error parsing template data", http.StatusBadRequest)
		log.Printf("Error parsing template data")
		return
	}

	err = template.ApplyTemplate(templateData)
	if err != nil {
		http.Error(w, "Error applying template", http.StatusInternalServerError)
		return
	}

	resultBytes, err := template.SaveAs()
	if err != nil {
		http.Error(w, "Error creating result document", http.StatusInternalServerError)
		return
	}

	// w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
	// w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", header.Filename))
	w.Write(resultBytes)
}

func main() {
	http.HandleFunc("/apply-template", handleTemplate)
	// http.ListenAndServe(":1137", nil)

	certFile := ""
	keyFile := ""

	fmt.Println("Server is running on 0137")
	err := http.ListenAndServeTLS(":1137", certFile, keyFile, nil)
	if err != nil {
		log.Fatal("ListenAndServeTLS: ", err)
	}
}
