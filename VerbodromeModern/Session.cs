using System;
using Microsoft.Data.Sqlite;

namespace VerbodromeModern
{
    public class Session
    {
        private readonly Random _random = new();
        private readonly string _dbPath;

        public ScoreBoard Score { get; } = new();

        private Session(string dbPath)
        {
            _dbPath = dbPath;
            using var con = new SqliteConnection($"Data Source={dbPath}");
            con.Open();
            Score.Load(con);
        }

        public static Session Load(string dbPath)
        {
            return new Session(dbPath);
        }

        public void Save()
        {
            using var con = new SqliteConnection($"Data Source={_dbPath}");
            con.Open();
            Score.Save(con);
        }

        public Question GenerateRandomQuestion(string[] verbs, string[] tenses, string[] persons)
        {
            var verb = verbs[_random.Next(verbs.Length)];
            var tense = ChooseTense(tenses);
            var person = persons[_random.Next(persons.Length)];
            return new Question(verb, tense, person);
        }

        private string ChooseTense(string[] tenses)
        {
            var stats = new (string Tense, double Accuracy)[tenses.Length];
            for (int i = 0; i < tenses.Length; i++)
            {
                var (c, t) = Score.StatsForTense(tenses[i]);
                double acc = t == 0 ? 0 : (double)c / t;
                stats[i] = (tenses[i], acc);
            }
            Array.Sort(stats, (a, b) => b.Accuracy.CompareTo(a.Accuracy));
            int sum = tenses.Length * (tenses.Length + 1) / 2;
            int pick = _random.Next(sum) + 1; // 1..sum
            int cumulative = 0;
            for (int i = 0; i < stats.Length; i++)
            {
                cumulative += i + 1;
                if (pick <= cumulative)
                    return stats[i].Tense;
            }
            return stats[^1].Tense;
        }

        public void RecordAnswer(Question question, bool correct)
        {
            Score.AddResult(question, correct);
        }
    }
}
