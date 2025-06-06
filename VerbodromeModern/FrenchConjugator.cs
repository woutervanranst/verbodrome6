using System;

namespace VerbodromeModern
{
    public class FrenchConjugator
    {
        private static readonly StringComparer IgnoreCase = StringComparer.OrdinalIgnoreCase;

        private static readonly Dictionary<string, Dictionary<string, Dictionary<string, string>>> Irregular = new(IgnoreCase)
        {
            ["être"] = new Dictionary<string, Dictionary<string, string>>(IgnoreCase)
            {
                ["Présent"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "suis",
                    ["tu"] = "es",
                    ["il"] = "est",
                    ["nous"] = "sommes",
                    ["vous"] = "êtes",
                    ["ils"] = "sont"
                },
                ["Passé composé"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "ai été",
                    ["tu"] = "as été",
                    ["il"] = "a été",
                    ["nous"] = "avons été",
                    ["vous"] = "avez été",
                    ["ils"] = "ont été"
                },
                ["Imparfait"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "étais",
                    ["tu"] = "étais",
                    ["il"] = "était",
                    ["nous"] = "étions",
                    ["vous"] = "étiez",
                    ["ils"] = "étaient"
                },
                ["Plus-que-parfait"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "avais été",
                    ["tu"] = "avais été",
                    ["il"] = "avait été",
                    ["nous"] = "avions été",
                    ["vous"] = "aviez été",
                    ["ils"] = "avaient été"
                },
                ["Passé simple"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "fus",
                    ["tu"] = "fus",
                    ["il"] = "fut",
                    ["nous"] = "fûmes",
                    ["vous"] = "fûtes",
                    ["ils"] = "furent"
                },
                ["Passé antérieur"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "eus été",
                    ["tu"] = "eus été",
                    ["il"] = "eut été",
                    ["nous"] = "eûmes été",
                    ["vous"] = "eûtes été",
                    ["ils"] = "eurent été"
                },
                ["Futur simple"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "serai",
                    ["tu"] = "seras",
                    ["il"] = "sera",
                    ["nous"] = "serons",
                    ["vous"] = "serez",
                    ["ils"] = "seront"
                },
                ["Futur antérieur"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "aurai été",
                    ["tu"] = "auras été",
                    ["il"] = "aura été",
                    ["nous"] = "aurons été",
                    ["vous"] = "aurez été",
                    ["ils"] = "auront été"
                },
                ["Subjonctif Présent"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "sois",
                    ["tu"] = "sois",
                    ["il"] = "soit",
                    ["nous"] = "soyons",
                    ["vous"] = "soyez",
                    ["ils"] = "soient"
                },
                ["Subjonctif Passé"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "aie été",
                    ["tu"] = "aies été",
                    ["il"] = "ait été",
                    ["nous"] = "ayons été",
                    ["vous"] = "ayez été",
                    ["ils"] = "aient été"
                },
                ["Subjonctif Imparfait"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "fusse",
                    ["tu"] = "fusses",
                    ["il"] = "fût",
                    ["nous"] = "fussions",
                    ["vous"] = "fussiez",
                    ["ils"] = "fussent"
                },
                ["Subjonctif Plus-que-parfait"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "eusse été",
                    ["tu"] = "eusses été",
                    ["il"] = "eût été",
                    ["nous"] = "eussions été",
                    ["vous"] = "eussiez été",
                    ["ils"] = "eussent été"
                },
                ["Conditionnel Présent"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "serais",
                    ["tu"] = "serais",
                    ["il"] = "serait",
                    ["nous"] = "serions",
                    ["vous"] = "seriez",
                    ["ils"] = "seraient"
                },
                ["Conditionnel Passé première forme"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "aurais été",
                    ["tu"] = "aurais été",
                    ["il"] = "aurait été",
                    ["nous"] = "aurions été",
                    ["vous"] = "auriez été",
                    ["ils"] = "auraient été"
                },
                ["Conditionnel Passé deuxième forme"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["je"] = "eusse été",
                    ["tu"] = "eusses été",
                    ["il"] = "eût été",
                    ["nous"] = "eussions été",
                    ["vous"] = "eussiez été",
                    ["ils"] = "eussent été"
                },
                ["Impératif Présent"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["tu"] = "sois",
                    ["nous"] = "soyons",
                    ["vous"] = "soyez"
                },
                ["Impératif Passé"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["tu"] = "aie été",
                    ["nous"] = "ayons été",
                    ["vous"] = "ayez été"
                },
                ["Participe Présent"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["*"] = "étant"
                },
                ["Participe Passé"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["*"] = "été"
                },
                ["Infinitif Présent"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["*"] = "être"
                },
                ["Infinitif Passé"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["*"] = "avoir été"
                },
                ["Gérondif Présent"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["*"] = "en étant"
                },
                ["Gérondif Passé"] = new Dictionary<string, string>(IgnoreCase)
                {
                    ["*"] = "en ayant été"
                }
            }
        };

        public string Conjugate(string infinitive, string tense, string person)
        {
            if (string.IsNullOrWhiteSpace(infinitive))
                throw new ArgumentException("Verb is required", nameof(infinitive));

            if (Irregular.TryGetValue(infinitive, out var tenses) && tenses.TryGetValue(tense, out var forms))
            {
                if (forms.TryGetValue(person, out var res))
                    return res;
                if (forms.TryGetValue("*", out var any))
                    return any;
                throw new ArgumentException("Invalid person for tense");
            }

            if (infinitive.EndsWith("er", StringComparison.OrdinalIgnoreCase))
                return ConjugateEr(infinitive, tense, person);
            if (infinitive.EndsWith("ir", StringComparison.OrdinalIgnoreCase))
                return ConjugateIr(infinitive, tense, person);
            if (infinitive.EndsWith("re", StringComparison.OrdinalIgnoreCase))
                return ConjugateRe(infinitive, tense, person);

            throw new ArgumentException("Unsupported verb");
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
