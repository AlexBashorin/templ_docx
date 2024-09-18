package paste

import (
	"archive/zip"
	"fmt"
	"image"
	"io"
	"strings"

	"github.com/boombuler/barcode"
	"github.com/boombuler/barcode/qr"
)

// DocxTemplate представляет структуру DOCX файла
type DocxTemplate struct {
	Files map[string]*zip.File
}

// TemplateData содержит данные для замены в шаблоне
type TemplateData struct {
	Text    map[string]string
	Images  map[string]ImageData
	QRCodes map[string]string
	Tables  map[string][][]string
}

// ImageData содержит информацию об изображении
type ImageData struct {
	Path   string
	Width  int
	Height int
}

// LoadTemplate загружает DOCX файл как шаблон
func LoadTemplate(path string) (*DocxTemplate, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	template := &DocxTemplate{
		Files: make(map[string]*zip.File),
	}

	for _, file := range reader.File {
		template.Files[file.Name] = file
	}

	return template, nil
}

// ApplyTemplate применяет данные к шаблону
func (t *DocxTemplate) ApplyTemplate(data TemplateData) error {
	// Обработка document.xml
	docXml, err := t.getXMLContent("word/document.xml")
	if err != nil {
		return err
	}

	// Замена текста
	for key, value := range data.Text {
		docXml = strings.ReplaceAll(docXml, "{{"+key+"}}", value)
	}

	// Замена изображений
	for key, imgData := range data.Images {
		placeholder := "{{" + key + "}}"
		replacement := fmt.Sprintf(`<w:drawing><wp:inline><wp:extent cx="%d" cy="%d"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId%s"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`, imgData.Width*9525, imgData.Height*9525, key)
		docXml = strings.ReplaceAll(docXml, placeholder, replacement)
	}

	// Замена QR-кодов
	for key, value := range data.QRCodes {
		placeholder := "{{" + key + "}}"
		qrCode, err := generateQRCode(value)
		if err != nil {
			return err
		}
		// Аналогично изображениям, но с данными QR-кода
		replacement := fmt.Sprintf(`<w:drawing><wp:inline><wp:extent cx="%d" cy="%d"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId%s"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`, 200*9525, 200*9525, key)
		docXml = strings.ReplaceAll(docXml, placeholder, replacement)
	}

	// Замена таблиц
	for key, tableData := range data.Tables {
		placeholder := "{{" + key + "}}"
		tableXml := generateTableXml(tableData)
		docXml = strings.ReplaceAll(docXml, placeholder, tableXml)
	}

	// Сохранение изменений
	t.Files["word/document.xml"].Data = []byte(docXml)

	return nil
}

// SaveAs сохраняет измененный DOCX файл
func (t *DocxTemplate) SaveAs(path string) error {
	// Реализация сохранения файла
	return nil
}

func (t *DocxTemplate) getXMLContent(fileName string) (string, error) {
	file, ok := t.Files[fileName]
	if !ok {
		return "", fmt.Errorf("file not found: %s", fileName)
	}

	reader, err := file.Open()
	if err != nil {
		return "", err
	}
	defer reader.Close()

	content, err := io.ReadAll(reader)
	if err != nil {
		return "", err
	}

	return string(content), nil
}

func generateQRCode(data string) (image.Image, error) {
	qrCode, err := qr.Encode(data, qr.M, qr.Auto)
	if err != nil {
		return nil, err
	}

	qrCode, err = barcode.Scale(qrCode, 200, 200)
	if err != nil {
		return nil, err
	}

	return qrCode, nil
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
