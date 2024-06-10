using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

class Program
{
    static readonly List<long> ActiveChatIds = new List<long>();
    static readonly Dictionary<long, DateTime> LastMessageTimes = new Dictionary<long, DateTime>();

    static readonly int MinMessageIntervalSeconds = 1;

    static string selectedCourse = null;
    static string spreadsheetId;
    static string selectedSheet = null;
    static string groupName;
    static TelegramBotClient botClient;
    static DateTime botStartTime;

    static async Task<List<string>> GetSpreadsheetSheets(string spreadsheetId)
    {
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

        var spreadsheet = await service.Spreadsheets.Get(spreadsheetId).ExecuteAsync();
        var sheets = spreadsheet.Sheets.Select(sheet => sheet.Properties.Title).ToList();
        return sheets;
    }

    static async Task HandlePollingErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
    {
        var apiRequestException = exception as ApiRequestException;
        if (apiRequestException != null && apiRequestException.ErrorCode == 403)
        {
            Console.WriteLine("Бот был удален из группы.");
        }
        else
        {
            Console.WriteLine($"Telegram API Error:\n[{apiRequestException?.ErrorCode}]\n{exception.Message}");
        }
    }

    static async Task Main(string[] args)
    {
        botClient = new TelegramBotClient("7086897352:AAGI-Fd_O5mnb-w4IR5W-fnwIGGvRpXmoRo");

        botStartTime = DateTime.UtcNow;

        using CancellationTokenSource cts = new();

        ReceiverOptions receiverOptions = new()
        {
            AllowedUpdates = Array.Empty<UpdateType>()
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

        cts.Cancel();
    }

    static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
    {
        if (update.Message is not { } message)
            return;

        var chatId = message.Chat.Id;

        // Проверяем, было ли это сообщение от пользователя.
        if (message.From == null)
        {
            Console.WriteLine("Не удалось получить информацию о пользователе.");
            return;
        }

        var userId = message.From.Id;

        // Проверяем, если пользователь уже взаимодействовал с ботом, используя его ID.
        if (!ActiveChatIds.Contains(userId))
        {
            ActiveChatIds.Add(userId);
        }

        if (message.Date.ToUniversalTime() < botStartTime)
        {
            Console.WriteLine($"Сообщение от пользователя {chatId} проигнорировано, так как было отправлено до запуска бота.");
            return;
        }

        if (LastMessageTimes.TryGetValue(chatId, out DateTime lastMessageTime))
        {
            if ((DateTime.UtcNow - lastMessageTime).TotalSeconds < MinMessageIntervalSeconds)
            {
                Console.WriteLine($"Сообщение от пользователя {chatId} проигнорировано из-за частого отправления.");
                return;
            }
        }

        Console.WriteLine($"Получено сообщение от пользователя {chatId}: {message.Text}");

        LastMessageTimes[chatId] = DateTime.UtcNow;

        if (message.Text == "/start")
        {
            var replyMarkup = new ReplyKeyboardMarkup(new[]
            {
            new[]
            {
                new KeyboardButton("Курс 1"),
                new KeyboardButton("Курс 2"),
            },
            new[]
            {
                new KeyboardButton("Курс 3"),
                new KeyboardButton("Курс 4"),
            }
        })
            {
                OneTimeKeyboard = true,
                ResizeKeyboard = true
            };

            await botClient.SendTextMessageAsync(
                chatId: chatId,
                text: "Выберите курс:",
                replyMarkup: replyMarkup,
                cancellationToken: cancellationToken
            );
        }
        else if (message.Text != null)
        {
            // Проверяем, является ли сообщение текстом "Курс n".
            string text = message.Text.ToLower();

            if (text.StartsWith("курс "))
            {
                int course;
                if (int.TryParse(text.Substring(5), out course) && course >= 1 && course <= 4)
                {
                    selectedCourse = $"курс {course}";
                    switch (course)
                    {
                        case 1:
                            spreadsheetId = "1fPHNrnMDighH92ZkGN3GcNfZ5abo2IU1a12adLBqXj8";
                            break;
                        case 2:
                            spreadsheetId = "13aLnayHCB53L48H4KedZkHYntdCApkb7rv0gFkx3XjY";
                            break;
                        case 3:
                            spreadsheetId = "1E-ieb-3K4GwIBGXgiiy-koA2AEtxF9InGRW80wtAO7A";
                            break;
                        case 4:
                            spreadsheetId = "1r541rAXYzV9FO14zaUNbpQ00CXkc1aAWIrLd-RuP7uQ";
                            break;
                    }

                    // Запускаем окно выбора группы после выбора курса.
                    await OfferSheetSelection(botClient, chatId, cancellationToken);
                }
            }
            else if (await IsSheetName(text))
            {
                selectedSheet = text;
                groupName = selectedSheet;

                var replyMarkup = new ReplyKeyboardMarkup(new[]
                {
                new[]
                {
                    new KeyboardButton("Понедельник"),
                    new KeyboardButton("Вторник"),
                },
                new[]
                {
                    new KeyboardButton("Среда"),
                    new KeyboardButton("Четверг"),
                },
                new[]
                {
                    new KeyboardButton("Пятница"),
                    new KeyboardButton("Суббота"),
                }
            })
                {
                    OneTimeKeyboard = true,
                    ResizeKeyboard = true
                };

                await botClient.SendTextMessageAsync(
                    chatId: chatId,
                    text: "Выберите день недели:",
                    replyMarkup: replyMarkup,
                    cancellationToken: cancellationToken
                );
            }
            else
            {
                string dayOfWeekInput = text;
                if (dayOfWeekInput == "понедельник" || dayOfWeekInput == "вторник" || dayOfWeekInput == "среда" ||
                    dayOfWeekInput == "четверг" || dayOfWeekInput == "пятница" || dayOfWeekInput == "суббота")
                {
                    var schedule = await GetScheduleFromGoogleSheets(groupName, dayOfWeekInput);

                    var hideKeyboard = new ReplyKeyboardMarkup(new KeyboardButton[][] { });

                    await botClient.SendTextMessageAsync(
                        chatId: chatId,
                        text: schedule,
                        replyMarkup: hideKeyboard,
                        cancellationToken: cancellationToken
                    );
                }
            }
        }
    }


    static async Task<bool> IsSheetName(string text)
    {
        var sheets = await GetSpreadsheetSheets(spreadsheetId);
        return sheets.Contains(text, StringComparer.OrdinalIgnoreCase);
    }

    static async Task OfferSheetSelection(ITelegramBotClient botClient, long chatId, CancellationToken cancellationToken)
    {
        var sheets = await GetSpreadsheetSheets(spreadsheetId);

        var buttons = sheets.Select(sheet =>
            new KeyboardButton[] { new KeyboardButton(sheet) }
        ).ToArray();

        var replyMarkup = new ReplyKeyboardMarkup(buttons)
        {
            OneTimeKeyboard = true,
            ResizeKeyboard = true
        };

        await botClient.SendTextMessageAsync(
            chatId: chatId,
            text: "Выберите Группу:",
            replyMarkup: replyMarkup,
            cancellationToken: cancellationToken
        );
    }

    static async Task<string> GetScheduleFromGoogleSheets(string groupName, string dayOfWeekInput)
    {
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

        DayOfWeek requestedDayOfWeek = ConvertRussianDayOfWeekToEnum(dayOfWeekInput);
        if (requestedDayOfWeek != DayOfWeek.Sunday)
        {
            DateTime currentDate = DateTime.Today;
            DayOfWeek currentDayOfWeek = currentDate.DayOfWeek;
            int daysUntilRequestedDay = ((int)requestedDayOfWeek - (int)currentDayOfWeek + 7) % 7;
            DateTime requestedDate = currentDate.AddDays(daysUntilRequestedDay);
            Console.WriteLine($"The requested date is: {requestedDate:dd.MM.yyyy}");

            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = await request.ExecuteAsync();
            var values = response.Values;

            if (values != null && values.Count > 0)
            {
                StringBuilder scheduleBuilder = new StringBuilder();
                bool found = false;

                for (int i = 0; i < values.Count; i++)
                {
                    var row = values[i];

                    Console.WriteLine($"Processing row {i}: {string.Join(", ", row)}");

                    if (row.Count >= 4 && row[0].ToString().ToLower() == dayOfWeekInput.ToLower() && DateTime.TryParseExact(row[1].ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime rowDate))
                    {
                        Console.WriteLine($"Checking date: {rowDate:dd.MM.yyyy} against {requestedDate:dd.MM.yyyy}");
                        if (rowDate.Date == requestedDate.Date)
                        {
                            scheduleBuilder.AppendLine($"Расписание на {rowDate:dd.MM.yyyy} ({dayOfWeekInput}):");

                            for (int j = i; j < Math.Min(values.Count, i + 8); j++)
                            {
                                var currentRow = values[j];

                                if (currentRow.Count >= 7)
                                {
                                    Console.WriteLine($"Row {j}: {string.Join(", ", currentRow)}");
                                    scheduleBuilder.AppendLine($"{currentRow[4]}:");
                                    scheduleBuilder.AppendLine($"Время: {currentRow[3]}");
                                    scheduleBuilder.AppendLine($"Преподаватель: {currentRow[5]}");
                                    scheduleBuilder.AppendLine($"Аудитория: {currentRow[6]}");
                                    scheduleBuilder.AppendLine();
                                    found = true;
                                }
                            }

                            break;
                        }
                    }
                }

                if (found)
                {
                    return scheduleBuilder.ToString();
                }
                else
                {
                    return $"Расписание на {requestedDate:dd.MM.yyyy} ({dayOfWeekInput}) не найдено.";
                }
            }
            else
            {
                return $"Расписание для {groupName} не найдено.";
            }
        }
        else
        {
            return "Неверный день недели.";
        }
    }

    static DayOfWeek ConvertRussianDayOfWeekToEnum(string dayOfWeek)
    {
        return dayOfWeek.ToLower() switch
        {
            "понедельник" => DayOfWeek.Monday,
            "вторник" => DayOfWeek.Tuesday,
            "среда" => DayOfWeek.Wednesday,
            "четверг" => DayOfWeek.Thursday,
            "пятница" => DayOfWeek.Friday,
            "суббота" => DayOfWeek.Saturday,
            "воскресенье" => DayOfWeek.Sunday,
            _ => throw new ArgumentException("Неверный день недели", nameof(dayOfWeek)),
        };
    }
}
