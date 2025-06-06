using System;

namespace VerbodromeModern
{
    public class FrenchConjugator
    {
        public string Conjugate(string infinitive, string tense, string person)
        {
            if (string.IsNullOrWhiteSpace(infinitive))
                throw new ArgumentException("Verb is required", nameof(infinitive));

            if (infinitive.EndsWith("er", StringComparison.OrdinalIgnoreCase))
                return ConjugateEr(infinitive, tense, person);
            if (infinitive.EndsWith("ir", StringComparison.OrdinalIgnoreCase))
                return ConjugateIr(infinitive, tense, person);
            if (infinitive.EndsWith("re", StringComparison.OrdinalIgnoreCase))
                return ConjugateRe(infinitive, tense, person);

            throw new ArgumentException("Unsupported verb ending. Only regular 'er', 'ir' and 're' verbs are handled.");
        }

        private static string ConjugateEr(string infinitive, string tense, string person)
        {
            string stem = infinitive[..^2];
            return tense switch
            {
                "Présent" => person switch
                {
                    "je" => stem + "e",
                    "tu" => stem + "es",
                    "il" => stem + "e",
                    "nous" => stem + "ons",
                    "vous" => stem + "ez",
                    "ils" => stem + "ent",
                    _ => throw new ArgumentException("Invalid person")
                },
                "Imparfait" => ConjugateStem(stem, person, "ais", "ais", "ait", "ions", "iez", "aient"),
                "Futur" => ConjugateStem(infinitive, person, "ai", "as", "a", "ons", "ez", "ont"),
                "Conditionnel" => ConjugateStem(infinitive, person, "ais", "ais", "ait", "ions", "iez", "aient"),
                "Subjonctif" => ConjugateStem(stem, person, "e", "es", "e", "ions", "iez", "ent"),
                _ => throw new ArgumentException("Invalid tense")
            };
        }

        private static string ConjugateIr(string infinitive, string tense, string person)
        {
            string baseStem = infinitive[..^2];
            string longStem = baseStem + "iss";
            return tense switch
            {
                "Présent" => person switch
                {
                    "je" => baseStem + "is",
                    "tu" => baseStem + "is",
                    "il" => baseStem + "it",
                    "nous" => longStem + "ons",
                    "vous" => longStem + "ez",
                    "ils" => longStem + "ent",
                    _ => throw new ArgumentException("Invalid person")
                },
                "Imparfait" => ConjugateStem(longStem, person, "ais", "ais", "ait", "ions", "iez", "aient"),
                "Futur" => ConjugateStem(infinitive, person, "ai", "as", "a", "ons", "ez", "ont"),
                "Conditionnel" => ConjugateStem(infinitive, person, "ais", "ais", "ait", "ions", "iez", "aient"),
                "Subjonctif" => ConjugateStem(longStem, person, "e", "es", "e", "ions", "iez", "ent"),
                _ => throw new ArgumentException("Invalid tense")
            };
        }

        private static string ConjugateRe(string infinitive, string tense, string person)
        {
            string stem = infinitive[..^2];
            return tense switch
            {
                "Présent" => person switch
                {
                    "je" => stem + "s",
                    "tu" => stem + "s",
                    "il" => stem,
                    "nous" => stem + "ons",
                    "vous" => stem + "ez",
                    "ils" => stem + "ent",
                    _ => throw new ArgumentException("Invalid person")
                },
                "Imparfait" => ConjugateStem(stem, person, "ais", "ais", "ait", "ions", "iez", "aient"),
                "Futur" => ConjugateStem(stem + "r", person, "ai", "as", "a", "ons", "ez", "ont"),
                "Conditionnel" => ConjugateStem(stem + "r", person, "ais", "ais", "ait", "ions", "iez", "aient"),
                "Subjonctif" => ConjugateStem(stem, person, "e", "es", "e", "ions", "iez", "ent"),
                _ => throw new ArgumentException("Invalid tense")
            };
        }

        private static string ConjugateStem(string stem, string person, string je, string tu, string il, string nous, string vous, string ils)
        {
            return person switch
            {
                "je" => stem + je,
                "tu" => stem + tu,
                "il" => stem + il,
                "nous" => stem + nous,
                "vous" => stem + vous,
                "ils" => stem + ils,
                _ => throw new ArgumentException("Invalid person")
            };
        }
    }
}
