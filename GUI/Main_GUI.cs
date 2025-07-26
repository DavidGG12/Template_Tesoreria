using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.Files;
using Template_Tesoreria.Models;

namespace Template_Tesoreria.GUI
{
    public class Main_GUI
    {
        private MenuOption_Model _menu;

        public Main_GUI() 
        {
            this._menu = new MenuOption_Model();
        }

        public string mainMenu()
        {
            var options = this._menu.getMenu();
            ConsoleKey key;
            var nmBank = "";
            var slOpt = ""; //Variable para guardar la opcion que se escogió
            var confirm = "";
            var id = 1;

            do
            {
                Console.Clear();

                Console.Title = "Template Tesoreria";
                Console.ForegroundColor = ConsoleColor.Cyan;

                Console.WriteLine("╔════════════════════════════════════════════════════╗");
                Console.WriteLine("║                 TEMPLATE  TESORERIA                ║");
                Console.WriteLine("║                                                    ║");
                Console.WriteLine("║  Por favor selecciona el banco de la siguiente     ║");
                Console.WriteLine("║  lista para continuar:                             ║");
                Console.WriteLine("╚════════════════════════════════════════════════════╝\n");

                Console.ResetColor();

                Console.WriteLine("Selecciona la compañía que deseas generar:\n");

                foreach (var option in options)
                {
                    if (id.ToString() == option.ID)
                    {
                        Console.BackgroundColor = ConsoleColor.Gray;
                        Console.ForegroundColor = ConsoleColor.Black;
                        slOpt = option.Option;
                    }
                    Console.WriteLine(option.Option);
                    Console.ResetColor();
                }

                key = Console.ReadKey(true).Key;

                var chsOpt = options.Find(x => x.Option.Contains(slOpt));

                if (key == ConsoleKey.UpArrow || key == ConsoleKey.W)
                    id = (id == 1) ? 1 : int.Parse(chsOpt.ID) - 1;
                else if (key == ConsoleKey.DownArrow || key == ConsoleKey.S)
                    id = (id == options.Count) ? options.Count : int.Parse(chsOpt.ID) + 1;

                if (key == ConsoleKey.Enter)
                {
                    nmBank = chsOpt.Value;
                    Console.Write($"\n¿Está seguro de querer trabajar con {nmBank}? [S/N]: ");
                    confirm = Console.ReadLine().Trim();
                    if (confirm.Equals("s", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.Clear();
                        break;
                    }
                    else
                        key = ConsoleKey.UpArrow;
                    Console.Clear();
                }

            } while (key != ConsoleKey.Enter);

            return nmBank;
        }
    
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
    }
}
