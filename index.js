import TelegramBot from 'node-telegram-bot-api';
import ExcelJS from 'exceljs';

const bot = new TelegramBot(process.env.TELEGRAM_BOT_TOKEN, { polling: true });

bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;
  bot.sendMessage(chatId, 'Привет! Отправь мне файл с таблицей в формате Excel или CSV.');
});

bot.on('document', async (msg) => {
  const chatId = msg.chat.id;
  const fileId = msg.document.file_id;
  const fileName = msg.document.file_name;

  if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.csv')) {
    bot.sendMessage(chatId, 'Пожалуйста, отправьте файл в формате Excel (.xlsx) или CSV.');
    return;
  }

  try {
    const fileStream = await bot.getFileStream(fileId);
    const inputWorkbook = new ExcelJS.Workbook();

    if (fileName.endsWith('.xlsx')) {
      await inputWorkbook.xlsx.read(fileStream);
    } else if (fileName.endsWith('.csv')) {
      await inputWorkbook.csv.read(fileStream);
    }

    const inputWorksheet = inputWorkbook.worksheets[0];
    const outputWorkbook = new ExcelJS.Workbook();
    const outputWorksheet = outputWorkbook.addWorksheet('Processed Data');

    // Заголовки для выходной таблицы
    outputWorksheet.columns = [
      { header: 'Номер в Avito', key: 'avitoNumber' },
      { header: 'Регион размещения', key: 'region' },
      { header: 'Город', key: 'city' },
      { header: 'Адрес', key: 'address' },
      { header: 'Категория', key: 'category' },
      { header: 'Подкатегория', key: 'subcategory' },
      { header: 'Параметр', key: 'parameter' },
      { header: 'Название объявления', key: 'adName' },
      { header: 'Дата первой публикации на Avito', key: 'firstPublicationDate' },
      { header: 'Дата снятия с публикации на Avito', key: 'withdrawalDate' },
      { header: 'Просмотров на Avito', key: 'viewsCount' },
      { header: 'Запросов контактов на Avito', key: 'contactsRequests' },
      { header: 'Стоимость размещения, руб', key: 'accomodationCost' },
      { header: 'Дополнительные сервисы, руб', key: 'otherServicesCost' },
      { header: 'Всего, руб', key: 'allCost' },
      { header: 'Сумма за вычетом бонусов, руб', key: 'costWitheoutBonuses' },
      { header: 'Добавлений в избранное на Avito', key: 'likesCount' },
      { header: 'Сотрудник', key: 'employee' },
    ];

    // Обходим каждую строку входной таблицы
    inputWorksheet.eachRow((inputRow, rowNumber) => {
      if (rowNumber === 1) { // Пропускаем заголовок
        return;
      }

      const inputData = {
        adNumber: inputRow.getCell('A').value,
        region: inputRow.getCell('B').value,
        city: inputRow.getCell('C').value,
        address: inputRow.getCell('D').value,
        category: inputRow.getCell('E').value,
        subcategory: inputRow.getCell('F').value,
        parameter: inputRow.getCell('G').value,
        adName: inputRow.getCell('H').value,
        firstPublicationDate: inputRow.getCell('J').value,
        withdrawalDate: inputRow.getCell('K').value,
        viewsCount: inputRow.getCell('M').value,
        contacts: inputRow.getCell('P').value,
        likesCount: inputRow.getCell('V').value,
        adExpences: inputRow.getCell('W').value,
        deletedBonuses: inputRow.getCell('X').value,
        placementAndActionsExpences: inputRow.getCell('Y').value,
        promotionExpences: inputRow.getCell('Z').value,
        otherExpences: inputRow.getCell('AA').value,
      };

      // Преобразуем данные
      const outputData = {
        avitoNumber: inputData.adNumber,
        region: inputData.region,
        city: inputData.city,
        address: inputData.address,
        category: inputData.category,
        subcategory: inputData.subcategory,
        parameter: inputData.parameter,
        adName: inputData.adName,
        firstPublicationDate: inputData.firstPublicationDate,
        withdrawalDate: inputData.withdrawalDate,
        viewsCount: inputData.viewsCount,
        contactsRequests: inputData.contacts,
        accomodationCost: inputData.placementAndActionsExpences,
        otherServicesCost: inputData.promotionExpences,
        allCost: inputData.adExpences,
        costWitheoutBonuses: inputData.adExpences - inputData.deletedBonuses,
        likesCount: inputData.likesCount,
        employee: 'Менеджер',
      };

      // Вставляем данные в выходную таблицу
      outputWorksheet.addRow(outputData);
    });

    const buffer = await outputWorkbook.xlsx.writeBuffer();

    await bot.sendDocument(chatId, Buffer.from(buffer), {}, {
      filename: `processed_${fileName}`,
      contentType: 'application/octet-stream',
    });

    console.log(`Succesfully processed received file with name: ${fileName}`);

    bot.sendMessage(chatId, 'Ваш файл успешно обработан!');
  } catch (error) {
    console.error('Ошибка при обработке файла:', error);
    bot.sendMessage(chatId, 'Произошла ошибка при обработке файла.');
  }
});