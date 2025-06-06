using System;
using Spectre.Console;

namespace VerbodromeModern
{
    public class Program
    {
        public static void Main()
        {
            var session = Session.Load("verbodrome.db");
            var conjugator = new FrenchConjugator();
            string[] verbs = { "parler", "aimer", "manger", "regarder", "finir", "choisir", "réussir", "vendre", "attendre", "perdre", "être" };
            string[] tenses =
            {
                "Présent",
                "Passé composé",
                "Imparfait",
                "Plus-que-parfait",
                "Passé simple",
                "Passé antérieur",
                "Futur simple",
                "Futur antérieur",
                "Subjonctif Présent",
                "Subjonctif Passé",
                "Subjonctif Imparfait",
                "Subjonctif Plus-que-parfait",
                "Conditionnel Présent",
                "Conditionnel Passé première forme",
                "Conditionnel Passé deuxième forme",
                "Impératif Présent",
                "Impératif Passé",
                "Participe Présent",
                "Participe Passé",
                "Infinitif Présent",
                "Infinitif Passé",
                "Gérondif Présent",
                "Gérondif Passé"
            };
            string[] persons = { "je", "tu", "il", "nous", "vous", "ils" };

            AnsiConsole.Write(new Rule("Verbodrome Quiz").RuleStyle("yellow"));
            int questionCount = 0;
            while (true)
            {
                var q = session.GenerateRandomQuestion(verbs, tenses, persons);
                var input = AnsiConsole.Ask<string>($"{q.Prompt} (type 'quit' to exit)");
                if (string.Equals(input, "quit", StringComparison.OrdinalIgnoreCase))
                    break;
                var expected = conjugator.Conjugate(q.Verb, q.Tense, q.Person);
                bool correct = string.Equals(input.Trim(), expected, StringComparison.OrdinalIgnoreCase);
                session.RecordAnswer(q, correct);
                questionCount++;
                if (correct)
                    AnsiConsole.MarkupLine("[green]Correct![/]");
                else
                    AnsiConsole.MarkupLine($"[red]Wrong, expected '{expected}'[/]");

                AnsiConsole.MarkupLine($"[yellow]Question:[/] {questionCount}  [yellow]Score:[/] {session.Score.Correct}/{session.Score.Total} ({session.Score.Percentage:F1}%)\n");
            }

            session.Save();
            AnsiConsole.MarkupLine("Goodbye!");
        }
    }
}
