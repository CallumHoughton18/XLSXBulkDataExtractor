using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Input;

namespace XLSXBulkDataExtractor.BL.MVVM_Extensions
{
    public class RelayCommand : ICommand
    {
        public delegate void ICommandOnExecute();
        public delegate bool ICommandOnCanExecute();

        private ICommandOnExecute _execute;
        private ICommandOnCanExecute _canExecute;

        public RelayCommand(ICommandOnExecute onExecuteMethod, ICommandOnCanExecute onCanExecuteMethod = null)
        {
            _execute = onExecuteMethod;
            _canExecute = onCanExecuteMethod;
        }


        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute?.Invoke() ?? true;
        }

        public void Execute(object parameter)
        {
            _execute?.Invoke();
        }
    }
}
