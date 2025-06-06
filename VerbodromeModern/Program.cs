using System;

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

            Console.WriteLine("Verbodrome Quiz (type 'quit' to exit)\n");
            while (true)
            {
                var q = session.GenerateRandomQuestion(verbs, tenses, persons);
                Console.Write(q.Prompt + ": ");
                var input = Console.ReadLine();
                if (input == null || input.Trim().Equals("quit", StringComparison.OrdinalIgnoreCase))
                    break;
                var expected = conjugator.Conjugate(q.Verb, q.Tense, q.Person);
                bool correct = string.Equals(input.Trim(), expected, StringComparison.OrdinalIgnoreCase);
                session.RecordAnswer(q, correct);
                Console.WriteLine(correct ? "Correct!" : $"Wrong, expected '{expected}'");
                Console.WriteLine($"Score: {session.Score.Correct}/{session.Score.Total} ({session.Score.Percentage:F1}%)\n");
            }
            session.Save();
            Console.WriteLine("Goodbye!");
        }
    }
}
