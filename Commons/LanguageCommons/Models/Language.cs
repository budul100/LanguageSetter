namespace LanguageCommons.Models
{
    public class Language
    {
        #region Public Properties

        public int Id { get; set; }

        public string Name { get; set; }

        #endregion Public Properties

        #region Public Methods

        public override string ToString()
        {
            return $"{Name} ({Id})";
        }

        #endregion Public Methods
    }
}