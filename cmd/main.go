package main

import (
	"bytes"
	"encoding/binary"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"templ/internal/paste"
)

func main() {

	http.HandleFunc("/generate_docx", getTempl)

	// certFile := ""
	// keyFile := ""

	fmt.Println("Server is running on 7131")
	http.ListenAndServe(":7131", nil)
	// err := http.ListenAndServeTLS(":7131", certFile, keyFile, nil)
	// if err != nil {
	// 	log.Fatal("ListenAndServeTLS: ", err)
	// }
}

func getTempl(w http.ResponseWriter, r *http.Request) {
	// get fromData
	err := r.ParseMultipartForm(10 << 20) // 10 MB limit
	if err != nil {
		http.Error(w, "Failed to parse form", http.StatusBadRequest)
		return
	}

	// Массив байт шаблона
	file, _, err := r.FormFile("template")
	if err != nil {
		http.Error(w, "Failed to get template file", http.StatusBadRequest)
		return
	}
	defer file.Close()
	templateBytes, err := io.ReadAll(file)
	if err != nil {
		http.Error(w, "Failed to read template file", http.StatusInternalServerError)
		return
	}

	// Текстовые переменные
	texts := r.FormValue("texts")
	if texts == "" {
		fmt.Println("No seneded texts")
	}
	var textVars map[string]string
	err = json.Unmarshal([]byte(texts), &textVars)
	if err != nil {
		http.Error(w, "Failed to parse data JSON", http.StatusBadRequest)
		return
	}
	fmt.Printf("texts: %v\n", textVars)

	// Изображения
	imgFiles, _, err := r.FormFile("images")
	if err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}
	defer imgFiles.Close()

	// Читаем все данные
	data, err := io.ReadAll(imgFiles)
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	// Создаем буфер для чтения данных
	buf := bytes.NewBuffer(data)

	imgData := make(map[string][]byte)

	// Читаем данные, пока буфер не пуст
	for buf.Len() > 0 {
		// Читаем длину имени
		var nameLen uint8
		if err := binary.Read(buf, binary.BigEndian, &nameLen); err != nil {
			http.Error(w, "Error reading name length", http.StatusBadRequest)
			return
		}

		// Читаем имя
		name := make([]byte, nameLen)
		if _, err := buf.Read(name); err != nil {
			http.Error(w, "Error reading name", http.StatusBadRequest)
			return
		}

		// Читаем оставшиеся данные как изображение
		imgData[string(name)] = buf.Bytes()
		buf.Reset()
	}

	// Ссылки
	links := r.FormValue("links")
	if links == "" {
		fmt.Println("No sended links")
	}
	var linkData map[string]string
	err = json.Unmarshal([]byte(links), &linkData)
	if err != nil {
		http.Error(w, "Failed to parse links JSON", http.StatusBadRequest)
		return
	}
	fmt.Printf("links: %v\n", linkData)

	// QR codes
	qrElma := r.FormValue("qrcode")
	if qrElma == "" {
		fmt.Println("No QR codes from ELMA")
	}
	var qrdata map[string]string
	err = json.Unmarshal([]byte(qrElma), &qrdata)
	if err != nil {
		http.Error(w, "Failed to get qrcode data JSON", http.StatusBadRequest)
		return
	}

	// Обработка
	result, err := paste.ReplaceVariablesInDOCX(templateBytes, textVars, imgData, linkData, qrdata)
	if err != nil {
		http.Error(w, "Failed parse: ReplaceVariablesInDOCX: error: "+err.Error(), http.StatusBadRequest)
		panic(err)
	}
	fmt.Printf("result: %v\n", "pasted")

	// запись файла
	// err = os.WriteFile("../result.docx", result, 0644)
	// if err != nil {
	// 	panic(err)
	// }

	_, errWrite := w.Write(result)
	if errWrite != nil {
		log.Fatal(errWrite)
	}
}

// type TemplateData struct {
// 	Variables map[string]string `json:"variables"`
// 	Images    map[string]string `json:"images"` // ключ - имя переменной, значение - имя файла
// }
// ----------- Пример использования
// docxContent, err := os.ReadFile("../template.docx")
// if err != nil {
// 	panic(err)
// }

// // Текстовые значения
// variables := map[string]string{
// 	"lol": "003",
// 	"kek": "004",
// }

// // Изображения
// idata, err := os.ReadFile("../pn.png")
// if err != nil {
// 	panic(err)
// }
// idataj, err := os.ReadFile("../jp.jpg")
// if err != nil {
// 	panic(err)
// }
// images := map[string][]byte{
// 	"myimg": idata,
// 	"secj":  idataj,
// }

// // Ссылки
// links := map[string]string{
// 	"cat": `https://www.apple.com/`,
// 	"dog": `https://google.com`,
// }
// ----------- END Пример использования

// func generateDOCX(w http.ResponseWriter, r *http.Request) {
// 	if r.Method != http.MethodPost {
// 		http.Error(w, "Only POST method is allowed", http.StatusMethodNotAllowed)
// 		return
// 	}

// 	// Parse multipart form
// 	err := r.ParseMultipartForm(10 << 20) // 10 MB limit
// 	if err != nil {
// 		http.Error(w, "Failed to parse form", http.StatusBadRequest)
// 		return
// 	}

// 	// Get template file
// 	file, header, err := r.FormFile("template")
// 	if err != nil {
// 		http.Error(w, "Failed to get template file", http.StatusBadRequest)
// 		return
// 	}
// 	defer file.Close()

// 	// Read template file into memory
// 	templateBytes, err := io.ReadAll(file)
// 	if err != nil {
// 		http.Error(w, "Failed to read template file", http.StatusInternalServerError)
// 		return
// 	}

// 	// Get data
// 	// dataStr := r.FormValue("data")
// 	// if dataStr == "" {
// 	// 	http.Error(w, "Data is missing", http.StatusBadRequest)
// 	// 	return
// 	// }

// 	// var data map[string]interface{}
// 	// err = json.Unmarshal([]byte(dataStr), &data)
// 	// if err != nil {
// 	// 	http.Error(w, "Failed to parse data JSON", http.StatusBadRequest)
// 	// 	return
// 	// }

// 	// Читаем JSON с данными шаблона
// 	var templateData TemplateData
// 	err = json.Unmarshal([]byte(r.FormValue("data")), &templateData)
// 	if err != nil {
// 		http.Error(w, "Invalid template data JSON", http.StatusBadRequest)
// 		return
// 	}

// 	// Process DOCX
// 	// resultBytes, err := processDocx(templateBytes, data)
// 	// if err != nil {
// 	// 	http.Error(w, fmt.Sprintf("Failed to process DOCX: %v", err), http.StatusInternalServerError)
// 	// 	return
// 	// }

// 	resultDocx, err := processTemplate(templateBytes, templateData, r)
// 	if err != nil {
// 		http.Error(w, fmt.Sprintf("Error processing DOCX: %v", err), http.StatusInternalServerError)
// 		return
// 	}

// 	fmt.Println(resultDocx)

// 	// Prepare response
// 	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", header.Filename))
// 	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

// 	// Send the processed DOCX
// 	_, err = w.Write(resultDocx)
// 	if err != nil {
// 		log.Printf("Failed to write response: %v", err)
// 	}
// }

// // HANDLE

// func processTemplate(docxContent []byte, data TemplateData, r *http.Request) ([]byte, error) {
// 	// Открываем DOCX как ZIP архив
// 	docxReader, err := zip.NewReader(bytes.NewReader(docxContent), int64(len(docxContent)))
// 	if err != nil {
// 		return nil, err
// 	}

// 	// Создаем новый ZIP архив для результата
// 	var resultBuffer bytes.Buffer
// 	resultZip := zip.NewWriter(&resultBuffer)

// 	imageIdCounter := 1000
// 	addedImages := make(map[string]string) // ключ - имя переменной, значение - rId

// 	for _, file := range docxReader.File {
// 		if file.Name == "word/document.xml" {
// 			// Обрабатываем основной документ
// 			doc, err := readZipFile(file)
// 			if err != nil {
// 				return nil, err
// 			}

// 			fmt.Println(doc)

// 			// Заменяем текстовые переменные
// 			for key, value := range data.Variables {
// 				// pattern := regexp.MustCompile(`@` + regexp.QuoteMeta(key))
// 				// doc = pattern.ReplaceAll(doc, []byte(value))
// 				// doc = strings.ReplaceAll(doc, "@"+key, value)
// 				// fmt.Printf("key %v", key)
// 				// doc = strings.ReplaceAll(doc, "@"+key, value)

// 				// pattern := regexp.MustCompile(`({?{)\s*` + regexp.QuoteMeta(key) + `\s*(}?})`)
// 				// doc = pattern.ReplaceAll(doc, []byte(value))

// 				pattern := regexp.MustCompile(`@` + regexp.QuoteMeta(key))
// 				doc = pattern.ReplaceAll(doc, []byte(value))
// 			}

// 			// Заменяем изображения
// 			for key, _ := range data.Images {
// 				if _, exists := addedImages[key]; !exists {
// 					imageIdCounter++
// 					rId := fmt.Sprintf("rId%d", imageIdCounter)
// 					addedImages[key] = rId
// 					patt := regexp.MustCompile(regexp.QuoteMeta(key))
// 					doc = patt.ReplaceAll(doc, []byte(createImageXml(rId)))
// 					// doc = strings.Replace(doc, key, createImageXml(rId), 1)
// 				}
// 			}

// 			err = writeZipFile(resultZip, file.Name, []byte(doc))
// 			if err != nil {
// 				return nil, err
// 			}
// 		} else if file.Name == "word/_rels/document.xml.rels" {
// 			// Добавляем связи с новыми изображениями
// 			rels, err := readZipFile(file)
// 			if err != nil {
// 				return nil, err
// 			}
// 			// image name
// 			for key, rId := range addedImages {
// 				fileName := data.Images[key]
// 				rels = strings.Replace(rels, "</Relationships>", createImageRelationship(rId, fileName)+"</Relationships>", 1)
// 			}
// 			err = writeZipFile(resultZip, file.Name, []byte(rels))
// 			if err != nil {
// 				return nil, err
// 			}
// 		} else if file.Name == "[Content_Types].xml" {
// 			// Добавляем информацию о типах содержимого для изображений
// 			types, err := readZipFile(file)
// 			if err != nil {
// 				return nil, err
// 			}
// 			types = []byte(addContentTypes(string(types)))
// 			err = writeZipFile(resultZip, file.Name, []byte(types))
// 			if err != nil {
// 				return nil, err
// 			}
// 		} else {
// 			// Копируем остальные файлы без изменений
// 			err = copyZipFile(resultZip, file)
// 			if err != nil {
// 				return nil, err
// 			}
// 		}
// 	}

// 	// Добавляем новые изображения
// 	for key, fileName := range data.Images {
// 		imageFile, _, err := r.FormFile(key)
// 		if err != nil {
// 			return nil, fmt.Errorf("image file %s is missing", key)
// 		}
// 		defer imageFile.Close()

// 		imageContent, err := io.ReadAll(imageFile)
// 		if err != nil {
// 			return nil, fmt.Errorf("unable to read image file %s", key)
// 		}

// 		err = writeZipFile(resultZip, "word/media/"+fileName, imageContent)
// 		if err != nil {
// 			return nil, err
// 		}
// 	}

// 	err = resultZip.Close()
// 	if err != nil {
// 		return nil, err
// 	}

// 	return resultBuffer.Bytes(), nil
// }

// func readZipFile(file *zip.File) ([]byte, error) {
// 	rc, err := file.Open()
// 	if err != nil {
// 		return nil, err
// 	}
// 	defer rc.Close()
// 	content, err := io.ReadAll(rc)
// 	if err != nil {
// 		return nil, err
// 	}
// 	return content, nil
// }

// func writeZipFile(zw *zip.Writer, name string, content []byte) error {
// 	w, err := zw.Create(name)
// 	if err != nil {
// 		return err
// 	}
// 	_, err = w.Write(content)
// 	return err
// }

// func copyZipFile(zw *zip.Writer, file *zip.File) error {
// 	rc, err := file.Open()
// 	if err != nil {
// 		return err
// 	}
// 	defer rc.Close()

// 	w, err := zw.Create(file.Name)
// 	if err != nil {
// 		return err
// 	}

// 	_, err = io.Copy(w, rc)
// 	return err
// }

// func createImageXml(imageId string) string {
// 	return fmt.Sprintf(`<w:p>
// 		<w:r>
// 			<w:drawing>
// 				<wp:inline>
// 					<wp:extent cx="5486400" cy="3657600"/>
// 					<wp:docPr id="1" name="Picture 1"/>
// 					<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
// 						<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
// 							<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
// 								<pic:nvPicPr>
// 									<pic:cNvPr id="1" name="Picture 1"/>
// 									<pic:cNvPicPr/>
// 								</pic:nvPicPr>
// 								<pic:blipFill>
// 									<a:blip r:embed="%s"/>
// 									<a:stretch>
// 										<a:fillRect/>
// 									</a:stretch>
// 								</pic:blipFill>
// 								<pic:spPr>
// 									<a:xfrm>
// 										<a:off x="0" y="0"/>
// 										<a:ext cx="5486400" cy="3657600"/>
// 									</a:xfrm>
// 									<a:prstGeom prst="rect">
// 										<a:avLst/>
// 									</a:prstGeom>
// 								</pic:spPr>
// 							</pic:pic>
// 						</a:graphicData>
// 					</a:graphic>
// 				</wp:inline>
// 			</w:drawing>
// 		</w:r>
// 	</w:p>`, imageId)
// }

// func createImageRelationship(imageId, imageName string) string {
// 	return fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, imageId, imageName)
// }

// func addContentTypes(types string) string {
// 	contentTypes := map[string]string{
// 		"png":  `<Default Extension="png" ContentType="image/png"/>`,
// 		"jpg":  `<Default Extension="jpg" ContentType="image/jpeg"/>`,
// 		"jpeg": `<Default Extension="jpeg" ContentType="image/jpeg"/>`,
// 		"gif":  `<Default Extension="gif" ContentType="image/gif"/>`,
// 	}

// 	for ext, contentType := range contentTypes {
// 		if !strings.Contains(types, fmt.Sprintf(`Extension="%s"`, ext)) {
// 			types = strings.Replace(types, "</Types>", contentType+"</Types>", 1)
// 		}
// 	}

// 	return types
// }
