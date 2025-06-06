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
            var tense = tenses[_random.Next(tenses.Length)];
            var person = persons[_random.Next(persons.Length)];
            return new Question(verb, tense, person);
        }

        public void RecordAnswer(Question question, bool correct)
        {
            Score.AddResult(question, correct);
        }
    }
}
