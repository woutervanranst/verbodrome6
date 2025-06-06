using System;

public class FrenchVerbConjugator
{
    public string Conjugate(string infinitive, string tense, string person)
    {
        if (infinitive.EndsWith("er"))
        {
            return ConjugateGroup1(infinitive, tense, person);
        }
        else if (infinitive.EndsWith("ir"))
        {
            return ConjugateGroup2(infinitive, tense, person);
        }
        else if (infinitive.EndsWith("re"))
        {
            return ConjugateGroup3(infinitive, tense, person);
        }
        else
        {
            throw new ArgumentException("Unsupported verb ending. Only regular verbs ending in 'er', 'ir', or 're' are supported.");
        }
    }

    private string ConjugateGroup1(string infinitive, string tense, string person)
    {
        string stem = infinitive.Substring(0, infinitive.Length - 2);
        switch (tense)
        {
            case "Présent":
                switch (person)
                {
                    case "je": return stem + "e";
                    case "tu": return stem + "es";
                    case "il": return stem + "e";
                    case "nous": return stem + "ons";
                    case "vous": return stem + "ez";
                    case "ils": return stem + "ent";
                    default: throw new ArgumentException("Invalid person.");
                }
            case "Imparfait":
                return ConjugateStem(stem, person, "ais", "ais", "ait", "ions", "iez", "aient");
            case "Futur":
                return ConjugateStem(infinitive, person, "ai", "as", "a", "ons", "ez", "ont");
            case "Conditionnel":
                return ConjugateStem(infinitive, person, "ais", "ais", "ait", "ions", "iez", "aient");
            case "Subjonctif":
                return ConjugateStem(stem, person, "e", "es", "e", "ions", "iez", "ent");
            default:
                throw new ArgumentException("Invalid tense.");
        }
    }

    private string ConjugateGroup2(string infinitive, string tense, string person)
    {
        string baseStem = infinitive.Substring(0, infinitive.Length - 2);
        string longStem = baseStem + "iss";

        switch (tense)
        {
            case "Présent":
                if (person == "je" || person == "tu" || person == "il")
                {
                    if (person == "je") return baseStem + "is";
                    else if (person == "tu") return baseStem + "is";
                    else return baseStem + "it";
                }
                else
                {
                    if (person == "nous") return longStem + "ons";
                    else if (person == "vous") return longStem + "ez";
                    else if (person == "ils") return longStem + "ent";
                    else throw new ArgumentException("Invalid person.");
                }
            case "Imparfait":
                return ConjugateStem(longStem, person, "ais", "ais", "ait", "ions", "iez", "aient");
            case "Futur":
                return ConjugateStem(infinitive, person, "ai", "as", "a", "ons", "ez", "ont");
            case "Conditionnel":
                return ConjugateStem(infinitive, person, "ais", "as", "a", "ons", "ez", "ont");
            case "Subjonctif":
                return ConjugateStem(longStem, person, "e", "es", "e", "ions", "iez", "ent");
            default:
                throw new ArgumentException("Invalid tense.");
        }
    }

    private string ConjugateGroup3(string infinitive, string tense, string person)
    {
        string stem = infinitive.Substring(0, infinitive.Length - 2);

        switch (tense)
        {
            case "Présent":
                if (person == "je") return stem + "s";
                else if (person == "tu") return stem + "s";
                else if (person == "il") return stem;
                else if (person == "nous") return stem + "ons";
                else if (person == "vous") return stem + "ez";
                else if (person == "ils") return stem + "ent";
                else throw new ArgumentException("Invalid person.");
            case "Imparfait":
                return ConjugateStem(stem, person, "ais", "ais", "ait", "ions", "iez", "aient");
            case "Futur":
                string futureStem = stem + "r";
                return ConjugateStem(futureStem, person, "ai", "as", "a", "ons", "ez", "ont");
            case "Conditionnel":
                string condStem = stem + "r";
                return ConjugateStem(condStem, person, "ais", "ais", "ait", "ions", "iez", "aient");
            case "Subjonctif":
                return ConjugateStem(stem, person, "e", "es", "e", "ions", "iez", "ent");
            default:
                throw new ArgumentException("Invalid tense.");
        }
    }

    private string ConjugateStem(string stem, string person, string jeEnding, string tuEnding, string ilEnding, 
                                string nousEnding, string vousEnding, string ilsEnding)
    {
        switch (person)
        {
            case "je": return stem + jeEnding;
            case "tu": return stem + tuEnding;
            case "il": return stem + ilEnding;
            case "nous": return stem + nousEnding;
            case "vous": return stem + vousEnding;
            case "ils": return stem + ilsEnding;
            default: throw new ArgumentException("Invalid person.");
        }
    }
}
