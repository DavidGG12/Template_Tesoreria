using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Models
{
    public class MenuOption_Model
    {
        private List<MenuOption_Model> _menu;
        public string ID { get; set; }
        public string Option { get; set; }
        public string Value { get; set; }
    
        public List<MenuOption_Model> getMenu()
        {
            this._menu = new List<MenuOption_Model>()
            {
                new MenuOption_Model() { ID = "1", Option = "1. - INBURSA", Value = "Inbursa" },
                new MenuOption_Model() { ID = "2", Option = "2. - HSBC", Value = "HSBC" },
                new MenuOption_Model() { ID = "3", Option = "3. - BANCOMER", Value = "Bancomer" },
                new MenuOption_Model() { ID = "4", Option = "4. - SCOTIABANK", Value = "Scotiabank" },
                new MenuOption_Model() { ID = "5", Option = "5. - CITIBANAMEX", Value = "Citibanamex" },
                new MenuOption_Model() { ID = "6", Option = "6. - SANTANDER", Value = "Santander" },
                new MenuOption_Model() { ID = "7", Option = "7. - BANORTE", Value = "Banorte" }
            };

            return this._menu;
        }

        public void addOption(MenuOption_Model option)
        {
            this._menu.Add(option);
        }
    }
}
