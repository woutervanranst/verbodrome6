using System;
using System.Windows;

namespace VerbodromeModern
{
    public partial class MainWindow : Window
    {
        private readonly Session _session;
        private readonly FrenchConjugator _conjugator = new();

        private readonly string[] _verbs = { "parler", "aimer", "finir", "choisir", "vendre" };
        private readonly string[] _tenses = { "Présent", "Imparfait", "Futur", "Conditionnel", "Subjonctif" };
        private readonly string[] _persons = { "je", "tu", "il", "nous", "vous", "ils" };

        private Question _current = null!;

        public MainWindow()
        {
            InitializeComponent();
            _session = Session.Load("verbodrome.db");
            NextQuestion();
            UpdateScore();
        }

        private void NextQuestion()
        {
            _current = _session.GenerateRandomQuestion(_verbs, _tenses, _persons);
            QuestionText.Text = _current.Prompt;
            AnswerBox.Text = string.Empty;
            AnswerBox.Focus();
        }

        private void CheckButton_Click(object sender, RoutedEventArgs e)
        {
            string answer = AnswerBox.Text.Trim();
            if (string.IsNullOrEmpty(answer))
                return;

            string correct = _conjugator.Conjugate(_current.Verb, _current.Tense, _current.Person);
            bool ok = string.Equals(answer, correct, StringComparison.OrdinalIgnoreCase);
            _session.RecordAnswer(_current, ok);
            Feedback.Text = ok ? "Correct!" : $"Wrong. Correct answer: {correct}";
            UpdateScore();
            NextQuestion();
        }

        private void UpdateScore()
        {
            ScoreText.Text = $"Score: {_session.Score.Correct}/{_session.Score.Total} ({_session.Score.Percentage:F1}%)";
        }

        protected override void OnClosed(EventArgs e)
        {
            _session.Save();
            base.OnClosed(e);
        }
    }
}
