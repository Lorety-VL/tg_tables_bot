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

  // Получаем заголовки входной таблицы
  const headerRow = inputWorksheet.getRow(1);
  const columnIndices = { //  В некоторых ячейках выгрузки пробелы другого типа, это решают регулярки
    adNumber: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Номер объявления'),
    region: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Регион размещения'),
    city: headerRow.values.findIndex((value) => value === 'Город'),
    address: headerRow.values.findIndex((value) => value === 'Адрес'),
    category: headerRow.values.findIndex((value) => value === 'Категория'),
    subcategory: headerRow.values.findIndex((value) => value === 'Подкатегория'),
    parameter: headerRow.values.findIndex((value) => value === 'Параметр'),
    adName: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Название объявления'),
    firstPublicationDate: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Дата первой публикации'),
    withdrawalDate: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Дата снятия с публикации'),
    employee: headerRow.values.findIndex((value) => value === 'Сотрудник'),
    viewsCount: headerRow.values.findIndex((value) => value === 'Просмотры'),
    contacts: headerRow.values.findIndex((value) => value === 'Контакты'),
    likesCount: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Добавили в избранное'),
    adExpences: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Расходы на объявления'),
    deletedBonuses: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Списано бонусов на объявления'),
    placementAndActionsExpences: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Расходы на размещение и целевые действия'),
    promotionExpences: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Расходы на продвижение'),
    otherExpences: headerRow.values.findIndex((value) => value?.replace(/\s+/g, ' ') === 'Остальные расходы'),
  };

  console.log(columnIndices);
  

  // Обходим каждую строку входной таблицы
  inputWorksheet.eachRow((inputRow, rowNumber) => {
    if (rowNumber === 1) { // Пропускаем заголовок
      return;
    }

    const inputData = {
      adNumber: inputRow.getCell(columnIndices.adNumber).value,
      region: inputRow.getCell(columnIndices.region).value,
      city: inputRow.getCell(columnIndices.city).value,
      address: inputRow.getCell(columnIndices.address).value,
      category: inputRow.getCell(columnIndices.category).value,
      subcategory: inputRow.getCell(columnIndices.subcategory).value,
      parameter: inputRow.getCell(columnIndices.parameter).value,
      adName: inputRow.getCell(columnIndices.adName).value,
      firstPublicationDate: inputRow.getCell(columnIndices.firstPublicationDate).value,
      withdrawalDate: inputRow.getCell(columnIndices.withdrawalDate).value,
      employee: inputRow.getCell(columnIndices.employee).value,
      viewsCount: inputRow.getCell(columnIndices.viewsCount).value,
      contacts: inputRow.getCell(columnIndices.contacts).value,
      likesCount: inputRow.getCell(columnIndices.likesCount).value,
      adExpences: inputRow.getCell(columnIndices.adExpences).value,
      deletedBonuses: inputRow.getCell(columnIndices.deletedBonuses).value,
      placementAndActionsExpences: inputRow.getCell(columnIndices.placementAndActionsExpences).value,
      promotionExpences: inputRow.getCell(columnIndices.promotionExpences).value,
      otherExpences: inputRow.getCell(columnIndices.otherExpences).value,
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