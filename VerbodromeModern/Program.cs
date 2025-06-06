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
            string[] verbs = { "parler", "finir", "vendre" };
            string[] tenses = { "Présent", "Imparfait", "Futur" };
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
