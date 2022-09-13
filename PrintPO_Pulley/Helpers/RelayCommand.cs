using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace BFDataCrawler.Helper
{
    class RelayCommand<T>:ICommand
    {
        readonly Predicate<T> _canExecute;
        readonly Action<T> _execute;
    
        public RelayCommand(Action<T> execute) : this(null, execute)
        {

        }
        public RelayCommand(Predicate<T> canExecute, Action<T> execute)
        {
            if (execute == null)
                throw new ArgumentException("execute");
            _canExecute = canExecute;
            _execute = execute;
        }
                       
        public bool CanExecute(object parameter)
        {
            return _canExecute == null ? true : _canExecute((T)parameter);
        }

        public void Execute(object paramter)
        {
            _execute((T)paramter);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}
