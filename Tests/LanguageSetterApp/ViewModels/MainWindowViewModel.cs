using Prism.Mvvm;

namespace LanguageSetterApp.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        #region Private Fields

        private string _title = "Prism Application";

        #endregion Private Fields

        #region Public Constructors

        public MainWindowViewModel()
        {
        }

        #endregion Public Constructors

        #region Public Properties

        public string Title
        {
            get { return _title; }
            set { SetProperty(ref _title, value); }
        }

        #endregion Public Properties
    }
}