using System;

namespace VerbodromeModern
{
    public record Question(string Verb, string Tense, string Person)
    {
        public string Prompt => $"Conjugate '{Verb}' in {Tense} ({Person})";
    }
}
