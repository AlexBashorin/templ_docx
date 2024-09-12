package paste

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"regexp"
	"strings"
)

// V 1
// func ReplaceVariablesInDOCX(docxFile []byte, variables map[string]string, images map[string][]byte) ([]byte, error) {
// 	reader, err := zip.NewReader(bytes.NewReader(docxFile), int64(len(docxFile)))
// 	if err != nil {
// 		return nil, err
// 	}

// 	var resultBuf bytes.Buffer
// 	writer := zip.NewWriter(&resultBuf)

// 	var documentXML []byte
// 	var relsXML []byte
// 	imageRelationships := make(map[string]string)

// 	for _, file := range reader.File {
// 		switch file.Name {
// 		case "word/document.xml":
// 			documentXML, err = readZipFile(file)
// 			if err != nil {
// 				return nil, err
// 			}
// 		case "word/_rels/document.xml.rels":
// 			relsXML, err = readZipFile(file)
// 			if err != nil {
// 				return nil, err
// 			}
// 		default:
// 			if err = copyZipFile(writer, file); err != nil {
// 				return nil, err
// 			}
// 		}
// 	}

// 	// Создание записи самого файла
// 	fmt.Println(len(images))
// 	for key, imageData := range images {
// 		// rId := fmt.Sprintf("rId%d", len(imageRelationships)+1)
// 		rId := "rId5"
// 		// fileName := fmt.Sprintf("image%d.png", len(imageRelationships)+1)
// 		imageRelationships["myimg"] = rId

// 		// Добавляем изображение в zip
// 		imgWriter, err := writer.Create(fmt.Sprintf("word/media/%s", key))
// 		if err != nil {
// 			return nil, err
// 		}
// 		if _, err = imgWriter.Write(imageData); err != nil {
// 			return nil, err
// 		}

// 		// Обновляем relationships
// 		relsXML = []byte(strings.Replace(string(relsXML), "</Relationships>", createImageRelationship(rId, key)+"</Relationships>", 1))
// 		fmt.Printf("relsXML: %v\n", string(relsXML))
// 	}

// 	// Заменяем переменные и добавляем изображения в document.xml
// 	for key, value := range variables {
// 		pattern := regexp.MustCompile(`@` + regexp.QuoteMeta(key))
// 		documentXML = pattern.ReplaceAll(documentXML, []byte(value))
// 	}

// 	for key, rId := range imageRelationships {

// 		pattern := regexp.MustCompile(`@` + regexp.QuoteMeta(key))

// 		replacement := []byte(fmt.Sprintf(`
//             <w:p>
//               <w:r>
//                 <w:drawing>
//                   <wp:inline distT="0" distB="0" distL="0" distR="0">
//                     <wp:extent cx="5486400" cy="3657600"/>
//                     <wp:effectExtent l="0" t="0" r="0" b="0"/>
//                     <wp:docPr id="1" name="Picture 1"/>
//                     <wp:cNvGraphicFramePr>
//                       <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
//                     </wp:cNvGraphicFramePr>
//                     <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
//                       <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
//                         <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
//                           <pic:nvPicPr>
//                             <pic:cNvPr id="0" name="Picture 1"/>
//                             <pic:cNvPicPr/>
//                           </pic:nvPicPr>
//                           <pic:blipFill>
//                             <a:blip r:embed="%s">
//                               <a:extLst>
//                                 <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
//                                   <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
//                                 </a:ext>
//                               </a:extLst>
//                             </a:blip>
//                             <a:stretch>
//                               <a:fillRect/>
//                             </a:stretch>
//                           </pic:blipFill>
//                           <pic:spPr>
//                             <a:xfrm>
//                               <a:off x="0" y="0"/>
//                               <a:ext cx="5486400" cy="3657600"/>
//                             </a:xfrm>
//                             <a:prstGeom prst="rect">
//                               <a:avLst/>
//                             </a:prstGeom>
//                           </pic:spPr>
//                         </pic:pic>
//                       </a:graphicData>
//                     </a:graphic>
//                   </wp:inline>
//                 </w:drawing>
//               </w:r>
//             </w:p>
//         `, rId))
// 		// replacement := []byte(fmt.Sprintf(`<w:pict><v:shape style="width:100pt;height:100pt"><v:imagedata r:id="%s"/></v:shape></w:pict>`, rId))
// 		// repl := []byte(fmt.Sprintf(`<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="%s" /></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`, rId))

// 		// documentXML = pattern.ReplaceAll(documentXML, replacement)
// 		documentXML = pattern.ReplaceAll(documentXML, bytes.TrimSpace(replacement))

// 	}

// 	// Записываем обновленные файлы
// 	if err = writeZipFile(writer, "word/document.xml", documentXML); err != nil {
// 		return nil, err
// 	}
// 	if err = writeZipFile(writer, "word/_rels/document.xml.rels", relsXML); err != nil {
// 		return nil, err
// 	}

// 	err = writer.Close()
// 	if err != nil {
// 		return nil, err
// 	}

// 	return resultBuf.Bytes(), nil
// }

// func createImageRelationship(rId, fileName string) string {
// 	return fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, rId, fileName)
// }

// func readZipFile(file *zip.File) ([]byte, error) {
// 	rc, err := file.Open()
// 	if err != nil {
// 		return nil, err
// 	}
// 	defer rc.Close()

// 	return ioutil.ReadAll(rc)
// }

// func writeZipFile(writer *zip.Writer, filename string, content []byte) error {
// 	w, err := writer.Create(filename)
// 	if err != nil {
// 		return err
// 	}

// 	_, err = w.Write(content)
// 	return err
// }

// func copyZipFile(writer *zip.Writer, file *zip.File) error {
// 	rc, err := file.Open()
// 	if err != nil {
// 		return err
// 	}
// 	defer rc.Close()

// 	w, err := writer.Create(file.Name)
// 	if err != nil {
// 		return err
// 	}

// 	_, err = io.Copy(w, rc)
// 	return err
// }

// V2

// func InsertImage(docxFile, imagePath string) error {
// 	// Открываем DOCX как zip архив
// 	reader, err := zip.OpenReader(docxFile)
// 	if err != nil {
// 		return err
// 	}
// 	defer reader.Close()

// 	// Создаем буфер для нового zip файла
// 	var buf bytes.Buffer
// 	writer := zip.NewWriter(&buf)

// 	// Читаем содержимое нового изображения
// 	imageData, err := ioutil.ReadFile(imagePath)
// 	if err != nil {
// 		return err
// 	}

// 	imageId := "rId1000" // Уникальный ID для нового изображения
// 	imageName := filepath.Base(imagePath)

// 	for _, file := range reader.File {
// 		if file.Name == "word/document.xml" {
// 			// Заменяем @myimg на XML-структуру изображения
// 			doc, _ := file.Open()
// 			content, _ := ioutil.ReadAll(doc)
// 			imageXml := `<w:p>
//                 <w:r>
//                     <w:drawing>
//                         <wp:inline>
//                             <wp:extent cx="5486400" cy="3657600"/>
//                             <wp:effectExtent l="0" t="0" r="0" b="0"/>
//                             <wp:docPr id="1" name="Picture 1"/>
//                             <wp:cNvGraphicFramePr>
//                                 <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
//                             </wp:cNvGraphicFramePr>
//                             <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
//                                 <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
//                                     <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
//                                         <pic:nvPicPr>
//                                             <pic:cNvPr id="0" name="Picture 1"/>
//                                             <pic:cNvPicPr/>
//                                         </pic:nvPicPr>
//                                         <pic:blipFill>
//                                             <a:blip r:embed="` + imageId + `">
//                                                 <a:extLst>
//                                                     <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
//                                                         <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
//                                                     </a:ext>
//                                                 </a:extLst>
//                                             </a:blip>
//                                             <a:stretch>
//                                                 <a:fillRect/>
//                                             </a:stretch>
//                                         </pic:blipFill>
//                                         <pic:spPr>
//                                             <a:xfrm>
//                                                 <a:off x="0" y="0"/>
//                                                 <a:ext cx="5486400" cy="3657600"/>
//                                             </a:xfrm>
//                                             <a:prstGeom prst="rect">
//                                                 <a:avLst/>
//                                             </a:prstGeom>
//                                         </pic:spPr>
//                                     </pic:pic>
//                                 </a:graphicData>
//                             </a:graphic>
//                         </wp:inline>
//                     </w:drawing>
//                 </w:r>
//             </w:p>`
// 			newContent := strings.Replace(string(content), "@myimg", imageXml, -1)
// 			_, _ = writer.Create(file.Name)
// 			_, _ = writer.Write([]byte(newContent))
// 		} else if file.Name == "word/_rels/document.xml.rels" {
// 			// Добавляем новую связь в document.xml.rels
// 			doc, _ := file.Open()
// 			content, _ := ioutil.ReadAll(doc)
// 			newRel := `<Relationship Id="` + imageId + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/` + imageName + `"/>`
// 			newContent := strings.Replace(string(content), "</Relationships>", newRel+"</Relationships>", 1)
// 			_ = writer.Create(file.Name)
// 			_, _ = writer.Write([]byte(newContent))
// 		} else {
// 			// Копируем остальные файлы без изменений
// 			dst, _ := writer.Create(file.Name)
// 			src, _ := file.Open()
// 			_, _ = io.Copy(dst, src)
// 		}
// 	}

// 	// Добавляем новое изображение
// 	imgWriter, _ := writer.Create("word/media/" + imageName)
// 	_, _ = imgWriter.Write(imageData)

// 	writer.Close()

// 	// Записываем измененный DOCX обратно в файл
// 	return ioutil.WriteFile(docxFile, buf.Bytes(), 0644)
// }

// V 3

func ReplaceVariablesInDOCX(docxFile []byte, variables map[string]string, images map[string][]byte, links map[string]string) ([]byte, error) {
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

	for _, file := range reader.File {
		switch file.Name {
		case "word/document.xml":
			documentXML, err = readZipFile(file)
			if err != nil {
				return nil, fmt.Errorf("error reading document.xml: %v", err)
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
		pattern := regexp.MustCompile(`<w:t>@` + regexp.QuoteMeta(key) + "</w:t>")
		replacement := generateImageXML(rId)
		documentXML = pattern.ReplaceAll(documentXML, replacement)
		log.Printf("Replaced @%s with image XML (rId: %s)", key, rId)
	}

	// texts
	log.Printf("Replacing text variables")
	for key, value := range variables {
		pattern := regexp.MustCompile(`@` + regexp.QuoteMeta(key))
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

	for key, rId := range linksRelationships {
		pattern := regexp.MustCompile("<w:r><w:rPr></w:rPr><w:t>@" + regexp.QuoteMeta(key) + "</w:t></w:r>")
		replacement := generateLinkXML(rId, key)
		fmt.Printf("generateLinkXML: %v\n", replacement)
		documentXML = pattern.ReplaceAll(documentXML, replacement)
	}

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
