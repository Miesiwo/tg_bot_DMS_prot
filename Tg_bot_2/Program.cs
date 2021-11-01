using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.ReplyMarkups;

namespace Tg_bot_2
{
    class Program
    {
        private static string token { get; } = "2020421106:AAG8XZjrNS-rlL9bMgElTCCjdB5_gtu9ru0";
        private static TelegramBotClient client;
            static List<string> repl = new List<string>();
        static void Main(string[] args)
        {
            




            client = new TelegramBotClient(token);
            client.StartReceiving();
            client.OnMessage += OnMessageHandler;
            Console.ReadLine();
            client.StopReceiving();

        }

        private static async void OnMessageHandler(object sender, MessageEventArgs e)
        {
            var msg = e.Message;
            if (msg.Text != null)
            {
                Console.WriteLine(msg.Text);
                
                switch (msg.Text)
                {
                    case "/start":
                        await client.SendTextMessageAsync(msg.Chat.Id, msg.Text, replyMarkup: GetButtons());
                        break;
                    case "Рапорт майно":
                        repl.Clear();
                        client.OnMessage -= OnMessageHandler;
                        await client.SendTextMessageAsync(msg.Chat.Id,"Воинская часть:", replyMarkup: new ForceReplyMarkup { Selective = true } );
                        client.OnMessage += OnReplyHandler;
                        break;
                }

            }

        }

        private static async void OnReplyHandler(object sender, MessageEventArgs e)
        {
            var msg = e.Message;

            if(msg.ReplyToMessage != null && msg.ReplyToMessage.Text.Contains("Воинская часть:"))
            {
                repl.Add(msg.Text);
                await client.SendTextMessageAsync(msg.Chat.Id, "Имя фамилия (Например: О.Петров):", replyMarkup: new ForceReplyMarkup { Selective = true });
            }
            else if(msg.ReplyToMessage != null && msg.ReplyToMessage.Text.Contains("Имя фамилия (Например: О.Петров):"))
            {
                repl.Add(msg.Text);
                await client.SendTextMessageAsync(msg.Chat.Id, "Звание ФИО:", replyMarkup: new ForceReplyMarkup { Selective = true });
            }
            else if(msg.ReplyToMessage != null && msg.ReplyToMessage.Text.Contains("Звание ФИО:"))
            {
                repl.Add(msg.Text);
                await client.SendTextMessageAsync(msg.Chat.Id, "Дата:", replyMarkup: new ForceReplyMarkup { Selective = true });
            }
            else if(msg.ReplyToMessage != null && msg.ReplyToMessage.Text.Contains("Дата:"))
            {
                repl.Add(msg.Text);
            }

            if(repl.Count == 4)
            {
                WordManager.Replacer(repl,0);
                
                using (var fileStream = File.OpenRead(WordManager.copyRaport))
                {
                    await client.SendDocumentAsync(msg.Chat.Id, fileStream);
                }
            }
        }

        private static IReplyMarkup GetButtons()
        {
            return new ReplyKeyboardMarkup
            {
                Keyboard = new List<List<KeyboardButton>>
                {
                    new List<KeyboardButton>{ new KeyboardButton { Text = "Рапорт майно" } },
                    new List<KeyboardButton>{ new KeyboardButton { Text = "Рапорт виплата"}
                }
            }
            };
        }
    }
}


