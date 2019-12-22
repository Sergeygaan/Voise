using System;

namespace VoiceControlCommon
{
    public static class ParsingText
    {
        public static string[] Parsing(string text)
        {
            return text.Split(new[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
        }
    }
}
