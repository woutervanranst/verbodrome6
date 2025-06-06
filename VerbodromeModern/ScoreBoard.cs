using System.Collections.Generic;
using System.Linq;
using Microsoft.Data.Sqlite;

namespace VerbodromeModern
{
    public class ScoreBoard
    {
        private readonly Dictionary<(string Verb, string Tense, string Person), (int Correct, int Total)> _stats = new();

        public void AddResult(Question q, bool correct)
        {
            var key = (q.Verb, q.Tense, q.Person);
            if (!_stats.TryGetValue(key, out var entry))
                entry = (0, 0);
            entry.Total++;
            if (correct) entry.Correct++;
            _stats[key] = entry;
        }

        public int Correct => _stats.Values.Sum(e => e.Correct);
        public int Total => _stats.Values.Sum(e => e.Total);
        public double Percentage => Total == 0 ? 0 : (double)Correct / Total * 100.0;

        public (int Correct, int Total) StatsForTense(string tense)
        {
            int correct = 0;
            int total = 0;
            foreach (var (key, value) in _stats)
            {
                if (key.Tense == tense)
                {
                    correct += value.Correct;
                    total += value.Total;
                }
            }
            return (correct, total);
        }

        public void Load(SqliteConnection connection)
        {
            using var cmd = connection.CreateCommand();
            cmd.CommandText = "CREATE TABLE IF NOT EXISTS Score (Verb TEXT, Tense TEXT, Person TEXT, Correct INTEGER, Total INTEGER, PRIMARY KEY (Verb, Tense, Person));";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "SELECT Verb, Tense, Person, Correct, Total FROM Score";
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var key = (reader.GetString(0), reader.GetString(1), reader.GetString(2));
                var entry = (reader.GetInt32(3), reader.GetInt32(4));
                _stats[key] = entry;
            }
        }

        public void Save(SqliteConnection connection)
        {
            using var tran = connection.BeginTransaction();
            foreach (var (key, value) in _stats)
            {
                using var cmd = connection.CreateCommand();
                cmd.Transaction = tran;
                cmd.CommandText = @"INSERT INTO Score (Verb, Tense, Person, Correct, Total)
                                    VALUES ($v,$t,$p,$c,$tt)
                                    ON CONFLICT(Verb,Tense,Person)
                                    DO UPDATE SET Correct=$c, Total=$tt;";
                cmd.Parameters.AddWithValue("$v", key.Verb);
                cmd.Parameters.AddWithValue("$t", key.Tense);
                cmd.Parameters.AddWithValue("$p", key.Person);
                cmd.Parameters.AddWithValue("$c", value.Correct);
                cmd.Parameters.AddWithValue("$tt", value.Total);
                cmd.ExecuteNonQuery();
            }
            tran.Commit();
        }
    }
}
