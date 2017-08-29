using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class PasswordBoxEx : DependencyObject
    {

        //private Dictionary<WeakReference<PasswordBox>

        static PasswordBoxEx()
        {
            EventManager.RegisterClassHandler(typeof(PasswordBox), PasswordBox.PasswordChangedEvent, new RoutedEventHandler(PasswordChangedEvent), false);
        }

        public static string GetPassword(DependencyObject obj)
        {
            return (string)obj.GetValue(PasswordProperty);
        }

        public static void SetPassword(DependencyObject obj, string value)
        {
            obj.SetValue(PasswordProperty, value);
        }

        public static readonly DependencyProperty PasswordProperty =
            DependencyProperty.RegisterAttached("Password", typeof(string), typeof(DependencyObject), new PropertyMetadata(String.Empty));


        private static void PasswordChangedEvent(object sender, RoutedEventArgs e)
        {
            var pb = sender as PasswordBox;
            pb.SetValue(PasswordProperty, pb.Password);
        }
    }
}
