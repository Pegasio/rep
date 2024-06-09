using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

class Program
{

    static async Task<Task> HandlePollingErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
    {
        var apiRequestException = exception as ApiRequestException;
        if (apiRequestException != null && apiRequestException.ErrorCode == 403)
        {
            // Бот был удален из группы, игнорируем эту ошибку
            Console.WriteLine("Бот был удален из группы.");
        }
        else
        {
            // Обработка остальных ошибок
            Console.WriteLine($"Telegram API Error:\n[{apiRequestException?.ErrorCode}]\n{exception.Message}");
        }

        // Продолжаем работу
        return Task.CompletedTask;
    }
    static async Task Main(string[] args)
    {
        var botClient = new TelegramBotClient("7086897352:AAGI-Fd_O5mnb-w4IR5W-fnwIGGvRpXmoRo");

        using CancellationTokenSource cts = new();

        // StartReceiving does not block the caller thread. Receiving is done on the ThreadPool.
        ReceiverOptions receiverOptions = new()
        {
            AllowedUpdates = Array.Empty<UpdateType>() // receive all update types except ChatMember related updates
        };

        botClient.StartReceiving(
            updateHandler: async (bot, update, cancellationToken) => await HandleUpdateAsync(botClient, update, cancellationToken),
            pollingErrorHandler: HandlePollingErrorAsync,
            receiverOptions: receiverOptions,
            cancellationToken: cts.Token
        );

        var me = await botClient.GetMeAsync();

        Console.WriteLine($"Start listening for @{me.Username}");
        Console.ReadLine();

        // Send cancellation request to stop bot
        cts.Cancel();
    }

    static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, System.Threading.CancellationToken cancellationToken)
    {
        if (update.Message is not { } message)
            return;

        if (message.Type == MessageType.ChatMemberLeft && message.LeftChatMember.Id == botClient.BotId)
        {
            // Бот был удален из чата, проигнорируем это событие
            Console.WriteLine($"Бот был удален из чата {message.Chat.Id}");
        }

        var chatId = message.Chat.Id;

        if (message.Text == "/start")
        {
            // Создаем клавиатуру с кнопками дней недели
            var replyMarkup = new ReplyKeyboardMarkup(new[]
            {
            new[]
            {
                new KeyboardButton("Понедельник"),
                new KeyboardButton("Вторник"),
                new KeyboardButton("Среда")
            },
            new[]
            {
                new KeyboardButton("Четверг"),
                new KeyboardButton("Пятница"),
                new KeyboardButton("Суббота")
            }
        });

            await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Выберите день недели:",
                replyMarkup: replyMarkup,
                cancellationToken: cancellationToken
            );
        }
        else if (message.Text != null)
        {
            string dayOfWeekInput = message.Text.ToLower(); // Приводим текст сообщения к нижнему регистру для удобства сравнения

// Проверяем, содержит ли сообщение ключевые слова (дни недели)
            if (dayOfWeekInput == "понедельник"  || dayOfWeekInput == "вторник" ||  dayOfWeekInput == "среда" ||
                dayOfWeekInput == "четверг"  || dayOfWeekInput == "пятница" ||  dayOfWeekInput == "суббота")
            {
                // Вызываем метод для получения расписания
                var schedule = await GetScheduleFromGoogleSheets("ДИС-204.121.Б", dayOfWeekInput);

                // Отправляем расписание пользователю
                await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: schedule,
                    cancellationToken: cancellationToken
                );
            }
        }
    }


    static async Task<string> GetScheduleFromGoogleSheets(string groupName, string dayOfWeekInput)
    {
        // Ваш код для получения расписания из Google Sheets
        string spreadsheetId = "1r541rAXYzV9FO14zaUNbpQ00CXkc1aAWIrLd-RuP7uQ";
        string range = $"{groupName}!A:G";

        GoogleCredential credential;
        using (var stream = new FileStream("key.json", FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream)
                .CreateScoped(new[] { SheetsService.Scope.SpreadsheetsReadonly });
        }

        var service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "Google Sheets API Sample",
        });

        // Вычисляем ближайший будущий день недели
        DayOfWeek requestedDayOfWeek = ConvertRussianDayOfWeekToEnum(dayOfWeekInput);
        if (requestedDayOfWeek != DayOfWeek.Sunday)
        {
            DateTime currentDate = DateTime.Today;
            DayOfWeek currentDayOfWeek = currentDate.DayOfWeek;
            int daysUntilRequestedDay = ((int)requestedDayOfWeek - (int)currentDayOfWeek + 7) % 7;
            DateTime requestedDate = currentDate.AddDays(daysUntilRequestedDay);
            Console.WriteLine($"The requested date is: {requestedDate:dd.MM.yyyy}");


            // Используем запрашиваемую дату вместо текущей даты
            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = await request.ExecuteAsync();
            var values = response.Values;

            if (values != null && values.Count > 0)
            {
                // Ваш код для обработки данных из Google Sheets
                StringBuilder scheduleBuilder = new StringBuilder();
                bool found = false;



                for (int i = 0; i < values.Count; i++)
                {
                    var row = values[i];

                    // Отладочная информация для каждой строки
                    Console.WriteLine($"Processing row {i}: {string.Join(", ", row)}");

                    if (row.Count >= 4 && row[0].ToString().ToLower() == dayOfWeekInput.ToLower() && DateTime.TryParseExact(row[1].ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime rowDate))
                    {
                        Console.WriteLine($"Checking date: {rowDate:dd.MM.yyyy} against {requestedDate:dd.MM.yyyy}");
                        if (rowDate.Date == requestedDate.Date)
                        {
                            // Найдена подходящая строка, обрабатываем ее
                            scheduleBuilder.AppendLine($"Расписание на {rowDate:dd.MM.yyyy} ({dayOfWeekInput}):");

                            for (int j = i; j < Math.Min(values.Count, i + 8); j++)
                            {
                                var currentRow = values[j];

                                if (currentRow.Count >= 7)
                                {
                                    Console.WriteLine($"nhin{j}: {string.Join(", ", currentRow)}");
                                    scheduleBuilder.AppendLine($"{currentRow[4]}:");
                                    scheduleBuilder.AppendLine($"Время: {currentRow[3]}");
                                    scheduleBuilder.AppendLine($"Аудитория: {currentRow[6]}");
                                    scheduleBuilder.AppendLine($"Преподаватель: {currentRow[5]}");
                                    scheduleBuilder.AppendLine();
                                    found = true;

                                }
                            }

                            break; // выход из цикла, так как мы нашли нужную дату и обработали ее
                        }
                    }
                }

                if (found)
                {
                    return scheduleBuilder.ToString();
                }
                else
                {
                    return $"No schedule found for {requestedDate:dd.MM.yyyy} ({dayOfWeekInput}).";
                }
            }
            else
            {
                return $"No schedule found for group {groupName}.";
            }



        }
        else
        {
            return "Invalid day of the week.";
        }
    }

    static DayOfWeek ConvertRussianDayOfWeekToEnum(string dayOfWeek)
    {
        switch (dayOfWeek.ToLower())
        {
            case "понедельник":
                return DayOfWeek.Monday;
            case "вторник":
                return DayOfWeek.Tuesday;
            case "среда":
                return DayOfWeek.Wednesday;
            case "четверг":
                return DayOfWeek.Thursday;
            case "пятница":
                return DayOfWeek.Friday;
            case "суббота":
                return DayOfWeek.Saturday;
            case "воскресенье":
                return DayOfWeek.Sunday;
            default:
                throw new ArgumentException("Invalid day of the week", nameof(dayOfWeek));
        }
    }










}
