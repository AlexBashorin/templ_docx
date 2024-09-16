package paste

import (
	"image"
	"image/color"
)

func GenerateQRCode(data string, size int) *image.RGBA {
	img := image.NewRGBA(image.Rect(0, 0, size, size))

	// Простая хеш-функция для преобразования строки в паттерн
	hash := 0
	for _, char := range data {
		hash = (hash*31 + int(char)) % 256
	}

	for y := 0; y < size; y++ {
		for x := 0; x < size; x++ {
			// Используем hash для создания паттерна
			if (x*y+hash)%7 < 3 {
				img.Set(x, y, color.Black)
			} else {
				img.Set(x, y, color.White)
			}
		}
	}

	return img
}
