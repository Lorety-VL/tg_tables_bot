import TelegramBot from 'node-telegram-bot-api';
import processDocument from './processDocument.js';

export default async (token) => {
  const bot = new TelegramBot(token, { polling: true });

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

      const buffer = await processDocument(fileStream, fileName);

      await bot.sendDocument(chatId, Buffer.from(buffer), {}, {
        filename: `processed_${fileName}`,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      console.log(`Successfully processed received file with name: ${fileName}`);

      bot.sendMessage(chatId, 'Ваш файл успешно обработан!');
    } catch (error) {
      console.error('Ошибка при обработке файла:', error);
      bot.sendMessage(chatId, 'Произошла ошибка при обработке файла.' + error);
    }
  });
};