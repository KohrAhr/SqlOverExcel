using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Lib.UI
{
    /// <summary>
    ///     Помошник для работы с окнами
    /// </summary>
    public static class WindowsUI
    {
        public const double CONST_OPACITY = 0.78;
        public const double CONST_OPACITY_NORMAL = 1;

        /// <summary>
        ///     Find any open window of such type and bring to front and also restore windows state.
        ///     <para>Otherwise create new window.</para>
        /// </summary>
        /// <typeparam name="T">
        ///     Window
        /// </typeparam>
        public static void ProceedWindow<T>(bool standAlone = false) where T : Window
        {
            Window window = WindowsUI.FindWindow<T>();

            if (window == null)
            {
                WindowsUI.ShowWindow<T>(standAlone);
            }
            else
            {
                window.BringIntoView();
                if (window.Focusable)
                {
                    window.Focus();
                }
                if (window.WindowState == WindowState.Minimized)
                {
                    window.WindowState = WindowState.Normal;
                }
            }
        }

        /// <summary>
        ///     Создать и показать окно не модально
        /// </summary>
        /// <typeparam name="T">
        ///     Тип окна
        /// </typeparam>
        public static void ShowWindow<T>(bool standAlone = false) where T : Window
        {
            Window parentWindow = null;

            if (!standAlone)
            {
                parentWindow = GetTopWindow();
            }

            T dlg = (T)Activator.CreateInstance(typeof(T), new object[] { });
            dlg.Owner = parentWindow; 
            dlg.Show();
        }

        /// <summary>
        ///     Создать и показать окно как диалоговое (модальное окно)
        /// </summary>
        /// <typeparam name="T">
        ///     Тип окна
        /// </typeparam>
        /// <returns>
        ///     Результат закрытия окна (Ok, Cancel)
        /// </returns>
        public static bool? ShowWindowDialog<T>() where T : Window
        {
            bool? result = false;

            Window parentWindow = GetTopWindow();

            T dlg = (T)Activator.CreateInstance(typeof(T), new object[] { });

            dlg.Owner = parentWindow;

            RunInOpacityMode(
                parentWindow, 
                () => { result = dlg.ShowDialog(); }
            );

            return result;
        }

        public static void RunWindowDialog(Action action)
        {
            RunInOpacityMode(
                GetTopWindow(), 
                () => { action.Invoke(); }
            );
        }

        /// <summary>
        ///     Создать и показать окно, которому передаётся один дополнительный параметр, как диалоговое
        /// </summary>
        /// <typeparam name="T">
        ///     Тип окна
        /// </typeparam>
        /// <param name="item">
        ///     Параметр
        /// </param>
        /// <returns>
        ///     Ссылка на окно, для возможности получить данные из окна после того как оператор завершенил работу с ним
        /// </returns>
        public static T ShowWindowDialogEx<T>(object item = null) where T : Window
        {
            Window parentWindow = GetTopWindow();

            T dialogWindow = (T)Activator.CreateInstance(typeof(T), item == null ? new object[] { } : new object[] { item });

            dialogWindow.Owner = parentWindow;

            RunInOpacityMode(parentWindow, 
                () => { dialogWindow.ShowDialog(); }
            );

            return dialogWindow;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parentWindow"></param>
        /// <param name="action"></param>
        private static void RunInOpacityMode(Window parentWindow, Action action)
        {
            double parentWindowOpacity = CONST_OPACITY_NORMAL;
            try
            {
                if (parentWindow != null)
                {
                    parentWindowOpacity = parentWindow.Opacity;
                    parentWindow.Opacity = CONST_OPACITY;
                }

                action.Invoke();
            }
            finally
            {
                if (parentWindow != null)
                {
                    parentWindow.Opacity = parentWindowOpacity;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private static Window GetTopWindow()
        {
            WindowCollection windowCollection = Application.Current.Windows;
            List<Window> realWindowCollection = new List<Window>();

            foreach (Window window in windowCollection)
            {
                if (!window.GetType().Name.Contains("AdornerLayerWindow"))
                {
                    realWindowCollection.Add(window);
                }
            }

            // Get window we need
            return realWindowCollection.Count() > 0 ? realWindowCollection.Last() : null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static Window FindWindow<T>()
        {
            Window result = null;

            WindowCollection windowCollection = Application.Current.Windows;
            List<Window> realWindowCollection = new List<Window>();

            foreach (Window window in windowCollection)
            {
                if (window.GetType() == typeof(T))
                {
                    result = window;
                    break;
                }
            }

            return result;
        }
    }
}
