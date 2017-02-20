using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace OfficeDevPnP.PartnerPack.Setup.Components
{
    public class ActionCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public ActionCommand(Action execute) : this(execute, null)
        {

        }

        public ActionCommand(Action execute, Func<bool> canExecute)
        {
            if (execute == null)
                throw new ArgumentNullException(nameof(execute));
            this._canExecute = canExecute;
            this._execute = execute;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object parameter)
        {
            _execute();
            RaiseCanExecuteChanged();
        }

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
