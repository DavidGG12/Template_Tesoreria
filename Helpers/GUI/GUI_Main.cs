using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.Files;
using Template_Tesoreria.Models;
using static Microsoft.IO.RecyclableMemoryStreamManager;

namespace Template_Tesoreria.Helpers.GUI
{
    public class GUI_Main
    {
        public GUI_Main() { }

        public void Spinner(string message, CancellationToken token)
        {
            char[] secuence = { '|', '/', '-', '\\' };
            int pos = 0;

            Console.ForegroundColor = ConsoleColor.Green;

            while (!token.IsCancellationRequested)
            {
                Console.Write($"\r{message} {secuence[pos]}");
                pos = (pos + 1) % secuence.Length;
                Thread.Sleep(100);
            }

            Console.ResetColor();

            Console.Write($"\r{new string(' ', Console.WindowWidth)}");
            Console.Write("\rTerminado\n");
        }

        public string centerMessage(string message)
        {
            try
            {
                var consoleWidth = Console.WindowWidth;
                var padding = (consoleWidth - message.Length) / 2;
                return message.PadLeft(message.Length + padding);
            }
            catch (Exception ex)
            {
                viewErrorMessage(ex.Message);
            }

            return null;
        }

        public string viewMenu(string title, List<MenuOption_Model> menu)
        {
            ConsoleKey key;
            var opt = "";
            var id = 1;

            do
            {
                Console.Clear();

                Console.Title = title;
                Console.ForegroundColor = ConsoleColor.Cyan;

                Console.WriteLine($"{centerMessage("╔════════════════════════════════════════════════════╗")}");
                Console.WriteLine($"{centerMessage("║                                                    ║")}");
                Console.WriteLine($"{centerMessage("║                 TEMPLATE  TESORERIA                ║")}");
                Console.WriteLine($"{centerMessage("║                                                    ║")}");
                Console.WriteLine($"{centerMessage("║  Por favor selecciona el banco de la siguiente     ║")}");
                Console.WriteLine($"{centerMessage("║  lista para continuar:                             ║")}");
                Console.WriteLine($"{centerMessage("╚════════════════════════════════════════════════════╝")}\n");

                Console.ResetColor();

                Console.WriteLine("Selecciona la compañía que deseas generar:\n");

                foreach (var option in menu)
                {
                    if (id.ToString() == option.ID)
                    {
                        Console.BackgroundColor = ConsoleColor.Gray;
                        Console.ForegroundColor = ConsoleColor.Black;
                        opt = option.Option;
                    }
                    Console.WriteLine(option.Option);
                    Console.ResetColor();
                }

                key = Console.ReadKey(true).Key;

                var chsOpt = menu.Find(x => x.Option.Contains(opt)); //Es la opción que se escogió y se guarda toda la informacón de dicha opción.

                //FEATURE:  Se puede checar que, aunque se deje presionada la tecla, aún así no baje y suba seguido.
                //          O sea, que se tenga que teclear por cada vez que quieres bajar.
                if (key == ConsoleKey.UpArrow || key == ConsoleKey.W)
                    id = (id == 1) ? 1 : int.Parse(chsOpt.ID) - 1;
                else if (key == ConsoleKey.DownArrow || key == ConsoleKey.S)
                    id = (id == menu.Count) ? menu.Count : int.Parse(chsOpt.ID) + 1;

                if(key == ConsoleKey.Enter)
                {
                    Console.Write($"\n¿Está seguro de querer trabajar con {chsOpt.Value}? [S/N]: ");
                    key = Console.ReadKey(true).Key;

                    if(key == ConsoleKey.S)
                    {
                        Console.Clear();
                        return chsOpt.Value;
                    }
                }
            }
            while (key != ConsoleKey.Enter);

            return null;
        }

        public void viewMessage(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($"\n{centerMessage(message)}\n\n");
            Console.ResetColor();
        }

        public void viewErrorMessage(string message)
        {

        }
    }
}
