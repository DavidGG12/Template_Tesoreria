using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Configuration;
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

        //public string[] getParagraph(string text)
        //{
        //    try
        //    {
        //        /* 
        //         * Restamos 10 porque son 4 de padding al encabezado en general, 2 de los bordes que se dibujan del menú
        //         * y vamos a hacer 4 puntos de padding de derecha a izquierda. 
        //         */
        //        var width = Console.WindowWidth - 10; 
        //        string[] paragraph;
        //        int[] potApart; 
        //        var numDiv = (int)Math.Round((double)text.Length / width); //Para saber en cuántas cadenas se va a dividir el texto

        //        foreach(var)

        //        foreach(var num in Enumerable.Range(0, numDiv))
        //        {
        //            var sentence = "";
                    
        //        }

        //        return paragraph;
        //    }
        //    catch(Exception ex)
        //    {
        //        viewErrorMessage(ex.Message);
        //        return null;
        //    }
        //}

        private void setHeader(string title, List<string> description)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            var width = Console.WindowWidth - 6;


            Console.WriteLine($"{centerMessage($"╔{new string('═', width)}╗")}");
            Console.WriteLine($"{centerMessage($"║{new string(' ', width)}║")}");
            Console.WriteLine($"{centerMessage($"║{new string(' ', (width - title.Length) / 2)}{title.ToUpper()}{new string(' ', (width - title.Length) / 2)}║")}");
            Console.WriteLine($"{centerMessage($"║{new string(' ', width)}║")}");

            foreach (var lines in description)
                Console.WriteLine($"{centerMessage($"║{new string(' ', (width - lines.Length) / 2)}{lines}{new string(' ', (width - lines.Length) / 2)}║")}");

            Console.WriteLine($"{centerMessage($"║{new string(' ', width)}║")}");
            Console.WriteLine($"{centerMessage($"╚{new string('═', width)}╝")}\n");

            Console.ResetColor();
        }

        public string viewMenu(string title, string description, List<MenuOption_Model> menu)
        {
            ConsoleKey key;
            var opt = "";
            var id = 1;

            do
            {
                Console.Clear();
                Console.Title = title;

                List<string> paragraph = new List<string> { "Por favor selecciona el banco de la siguiente ", "lista para continuar: " };

                setHeader(title, paragraph);

                foreach (var option in menu)
                {
                    if (id.ToString() == option.ID)
                    {
                        Console.BackgroundColor = ConsoleColor.Gray;
                        Console.ForegroundColor = ConsoleColor.Black;
                        opt = option.Option;
                    }
                    Console.WriteLine($"{new string(' ', 4)}{option.Option}");
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
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write($"\n{message}\n\n");
            Console.ResetColor();
        }
    }
}
