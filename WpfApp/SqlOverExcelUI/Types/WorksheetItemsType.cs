using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Collections.ObjectModel;

namespace SqlOverExcelUI.Types
{
    public class WorksheetItemsType : ObservableCollection<WorksheetItemType>
    {
        private ObservableCollection<WorksheetItemType> items;

        /// <summary>
        ///     Constructor
        /// </summary>
        public WorksheetItemsType()
        {
            items = new ObservableCollection<WorksheetItemType>();
        }

        public new IEnumerable Items
        {
            get { return items; }
        }
    }

}
