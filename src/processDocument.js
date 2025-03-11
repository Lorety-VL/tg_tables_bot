import ExcelJS from 'exceljs';

export default async (fileStream, fileName) => {
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
      employee: inputRow.getCell('L').value,
      viewsCount: inputRow.getCell('N').value,
      contacts: inputRow.getCell('Q').value,
      likesCount: inputRow.getCell('W').value,
      adExpences: inputRow.getCell('X').value,
      deletedBonuses: inputRow.getCell('Y').value,
      placementAndActionsExpences: inputRow.getCell('Z').value,
      promotionExpences: inputRow.getCell('AA').value,
      otherExpences: inputRow.getCell('AB').value,
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
      employee: inputData.employee,
    };

    // Вставляем данные в выходную таблицу
    outputWorksheet.addRow(outputData);
  });

  const buffer = await outputWorkbook.xlsx.writeBuffer();

  return buffer;
};