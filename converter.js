const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const ExcelJS = require('exceljs');


// Путь к Excel-файлу
const excelFilePath = '/Users/marinasemenova/Downloads/прайсы.xlsx';

// Папка для сохранения изображений
const imagesFolder = './products_images';

// Создаем папку для изображений, если она не существует
if (!fs.existsSync(imagesFolder)) {
    fs.mkdirSync(imagesFolder);
}

// Читаем Excel с помощью xlsx (для работы с таблицей)
// const workbookXlsx = xlsx.readFile(excelFilePath);
// const sheetName = workbookXlsx.SheetNames[0];
// const sheet = workbookXlsx.Sheets[sheetName];
// const data = xlsx.utils.sheet_to_json(sheet);
let data = []
const workbookXlsx = xlsx.readFile(excelFilePath);

workbookXlsx.SheetNames.forEach(sheetName => {
    const sheet = workbookXlsx.Sheets[sheetName];
    const sheetData = xlsx.utils.sheet_to_json(sheet); // Получаем данные с листа
    data.push(...sheetData); // Добавляем в общий массив
    // console.log(`Данные с листа "${sheetName}":`, data);
});


console.log('первоначальная дата',data);

//словарь
const keysMapping = {
    '1d': 'id',
    'ррц': 'price',
    'Model': 'model',
    'описание': 'description',
    'форма': 'form',
    'разрешение': 'resolution',
    'матрица': 'matrix',
    'объектив': 'lens',
    'захватлиц': 'face_capture',
    'аудиоиo': 'audio_io',
    'тревожныйio': 'alarm_io',
    'статус': 'status',
    'ик': 'IR_range'
};
// Функция нормализации ключа: приводит к нижнему регистру, убирает пробелы и знаки препинания
function normalizeKey(key) {
    return key.toLowerCase().replace(/\s+/g, '').replace(/[^a-zа-я0-9]/gi, '');
}
// Функция расчёта расстояния Левенштейна для двух строк
function levenshtein(a, b) {
    const matrix = [];

    // инициализация первой строки и столбца
    for (let i = 0; i <= b.length; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= a.length; j++) {
        matrix[0][j] = j;
    }

    // заполнение матрицы
    for (let i = 1; i <= b.length; i++) {
        for (let j = 1; j <= a.length; j++) {
            if (b.charAt(i - 1) === a.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1, // замена символа
                    matrix[i][j - 1] + 1,     // вставка символа
                    matrix[i - 1][j] + 1      // удаление символа
                );
            }
        }
    }
    return matrix[b.length][a.length];
}

// Функция, которая получает новый ключ, учитывая возможные опечатки (порог = 1)
function getMappedKey(originalKey) {
    const normalized = normalizeKey(originalKey);
    if (keysMapping[normalized]) {
        return keysMapping[normalized];
    }
    // Пытаемся найти похожее название ключа с небольшим расстоянием
    let bestMatch = null;
    let bestDistance = Infinity;
    for (const key in keysMapping) {
        const dist = levenshtein(normalized, key);
        if (dist < bestDistance) {
            bestDistance = dist;
            bestMatch = key;
        }
    }
    if (bestDistance <= 1) { // можно подстроить порог при необходимости
        return keysMapping[bestMatch];
    }
    // Если не найдено – возвращаем исходный ключ без изменений
    return originalKey;
}

// Преобразование каждого объекта из data: заменяем ключи согласно сопоставлению
const products = data.map(item => {
    const newItem = {};
    for (const key in item) {
        const newKey = getMappedKey(key);
        newItem[newKey] = item[key];
    }
    return newItem;
});

console.log('Преобразованные продукты в массиве:', products);

//сохраняем в файл:
fs.writeFile('products.json', JSON.stringify(products, null, 2), (err) => {
    if (err) {
        console.error('Ошибка при сохранении файла:', err);
    } else {
        console.log('Файл products.json успешно сохранён!');
    }
});


// Функция для сохранения встроенного изображения
function saveImage(buffer, imageName) {
    const imagePath = path.join(imagesFolder, imageName);
    fs.writeFileSync(imagePath, buffer);
    console.log(`Изображение сохранено: ${imagePath}`);
}

// Функция для обработки встроенных изображений для одной страницыы
async function extractEmbeddedImagesForOnePage() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    const worksheet = workbook.worksheets[0];

    worksheet.getImages().forEach((img) => {
          const rowNumber = img.range.tl.nativeRow + 1; // Excel считает с 1
        // const rowNumber = img.range.tl.nativeRow - 1; // вариант с https://iampatelajeet.hashnode.dev/extracting-images-from-excel-files-in-nodejs , так хотя бы считает со втрой строчки экселя, но тогда минус 3 строчки в конце
        // const rowNumber = img.range.tl.nativeRow
         console.log('строка rowNumber',rowNumber);
        const row = worksheet.getRow(rowNumber);
        //  console.log('строка',row);
        
        
        const id = row.getCell(1).text.trim(); // ID из первой колонки (столбец A)

        if (id && img.imageId) {
            const image = workbook.model.media.find(m => m.index === img.imageId);
            if (image) {
                saveImage(image.buffer, `${id}.jpg`);
            }
        }
    });

    console.log('Извлечение встроенных изображений завершено.');
}
// Функция для обработки встроенных изображений для всех страниц с книги
async function extractEmbeddedImages() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);

    workbook.worksheets.forEach((worksheet, sheetIndex) => {
        worksheet.getImages().forEach((img, imgIndex) => {
            const rowNumber = img.range.tl.nativeRow + 1; // Excel считает с 1
            console.log(`Лист: ${worksheet.name}, строка: ${rowNumber}`);

            const row = worksheet.getRow(rowNumber);
            const id = row.getCell(1).text.trim(); // ID из первой колонки (столбец A)

            if (id && img.imageId) {
                const image = workbook.model.media.find(m => m.index === img.imageId);
                if (image) {
                    // Сохраняем изображение с учетом названия листа
                    saveImage(image.buffer, `${id}.jpg`);
                }
            }
        });
    });

    console.log('Извлечение встроенных изображений завершено.');
}


// Запускаем извлечение встроенных изображений
extractEmbeddedImages().catch(console.error);
// extractEmbeddedImagesForOnePage().catch(console.error)

console.log('Обработка завершена.');
