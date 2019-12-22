using Microsoft.Speech.Recognition;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Windows.Forms;

namespace VoiceControlRecognition
{
    /// <summary>
    /// Класс для распознания речи
    /// </summary>
    public class Recognition
    {
        #region TextLog

        private static string _textLog = string.Empty;

        public static void OnPropertyTextLog(PropertyChangedEventArgs e)
        {
            PropertyTextLog?.Invoke(null, e);
        }

        public static void OnPropertyTextLog(string propertyTextLog)
        {
            OnPropertyTextLog(new PropertyChangedEventArgs(propertyTextLog));
        }

        /// <summary>
        /// Переменная, хранящая значение денег персонажа
        /// </summary>
        public static string TextLog
        {
            get { return _textLog; }
            set
            {
                if (value != _textLog)
                {
                    _textLog = value;
                    OnPropertyTextLog(_textLog);
                }
            }
        }

        public static event PropertyChangedEventHandler PropertyTextLog;

        #endregion

        #region Command

        private static string _command = string.Empty;

        public static void OnPropertyCommand(PropertyChangedEventArgs e)
        {
            PropertyCommand?.Invoke(null, e);
        }

        public static void OnPropertyCommand(string propertyCommand)
        {
            OnPropertyCommand(new PropertyChangedEventArgs(propertyCommand));
        }

        /// <summary>
        /// Переменная, хранящая значение денег персонажа
        /// </summary>
        public static string Command
        {
            get { return _command; }
            set
            {
                if (value != _command)
                {
                    _command = value;
                    OnPropertyCommand(_command);
                }
            }
        }

        public static event PropertyChangedEventHandler PropertyCommand;

        #endregion

        private CultureInfo _culture;
        private SpeechRecognitionEngine _sre;

        /// <summary>
        /// Загруженные команды из файла
        /// </summary>
        private LoadCommand _loadCommand;


        private List<LoadArrayCommands> _loadArrayCommands = new List<LoadArrayCommands>();

        public Recognition(string cultureInfo)
        {
            _loadCommand = new LoadCommand();

            LoadCommands();

            StartGrammar(cultureInfo);
        }

        private void LoadCommands()
        {
            _loadArrayCommands = _loadCommand.OpenRead();
        }

        /// <summary>
        /// Старт распознания
        /// </summary>
        private void StartGrammar(string cultureInfo)
        {
            try
            {
                _culture = new CultureInfo(cultureInfo);
                _sre = new SpeechRecognitionEngine(_culture);

                // Setup event handlers
                _sre.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(SpeechDetected);
                _sre.RecognizeCompleted += new EventHandler<RecognizeCompletedEventArgs>(RecognizeCompleted);
                _sre.SpeechHypothesized += new EventHandler<SpeechHypothesizedEventArgs>(SpeechHypothesized);
                _sre.SpeechRecognitionRejected += new EventHandler<SpeechRecognitionRejectedEventArgs>(SpeechRecognitionRejected);
                _sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognized);

                // select input source
                _sre.SetInputToDefaultAudioDevice();

                // load grammar
                LoadGrammar();

                // start recognition
                _sre.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        /// <summary>
        /// Создание грамматики для распознавания речи
        /// </summary>
        /// <param name="semanticResult"></param>
        /// <returns></returns>
        private Choices CreateSample(List<Tuple<string, string>> semanticResult)
        {
            var grammarBuilders = new GrammarBuilder[semanticResult.Count];

            int index = 0;

            foreach (var currentSemantic in semanticResult)
            {
                grammarBuilders[index] = new SemanticResultValue(currentSemantic.Item1, currentSemantic.Item2);

                index++;
            }

            return new Choices(grammarBuilders);
        }

        /// <summary>
        /// Загрузка грамматики
        /// </summary>
        private void LoadGrammar()
        {
            //Загрузка команд из списка
            foreach (var currentCommand in _loadArrayCommands)
            {
                _sre.LoadGrammar(CreateGrammar(currentCommand.CommandTextReturn(), CreateSample(currentCommand.SemanticResultReturn())));
            }
        }

        private Grammar CreateGrammar(string commandText, Choices semanticResult)
        {
            var grammarBuilder = new GrammarBuilder(commandText, SubsetMatchingMode.SubsequenceContentRequired)
            {
                Culture = _culture
            };

            grammarBuilder.Append(new SemanticResultKey(commandText, semanticResult));

            return new Grammar(grammarBuilder);
        }

        private void SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            TextLog = $"[Speech Recognition Rejected: {e.Result.Text}][WhiteSmoke]";
        }

        private void SpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            TextLog = $"[Speech Hypothesized: {e.Result.Text} ({ e.Result.Confidence})][WhiteSmoke]";
        }

        private void RecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
        {
            TextLog = $"[Recognize Completed {e.Result.Text}][WhiteSmoke]";
        }

        private void SpeechDetected(object sender, SpeechDetectedEventArgs e)
        {
            TextLog = $"[Speech Detected: audio pos {e.AudioPosition}][WhiteSmoke]";
        }

        //Определение корректности текста
        private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            TextLog = $"[\tSpeech Recognized:][WhiteSmoke]";

            if (e.Result.Confidence < 0.35f)
            {
                TextLog = $"[{e.Result.Text} ({e.Result.Confidence})][WhiteSmoke]";
                return;
            }
            for (var i = 0; i < e.Result.Alternates.Count; ++i)
            {
                TextLog = $"[\tAlternate: {e.Result.Alternates[i].Text} ({e.Result.Alternates[i].Confidence})][WhiteSmoke]";
            }

            for (var i = 0; i < e.Result.Words.Count; ++i)
            {
                if (e.Result.Words[i].Confidence < 0.2f)
                {
                    TextLog = $"[\tWord: {e.Result.Words[i].Text} ({e.Result.Words[i].Confidence})][LightCoral]";
                    return;
                }
                else
                {
                    TextLog = $"[\tWord: {e.Result.Words[i].Text} ({e.Result.Words[i].Confidence})][WhiteSmoke]";
                }
            }

            TextLog = $"[{e.Result.Text} ({ e.Result.Confidence})][YellowGreen]";

            TextLog = $"[----------------------------------------------------------------][WhiteSmoke]";

            foreach (var s in e.Result.Semantics)
            {
                var number = s.Value.Value;

                for (int index = 0; index < _loadArrayCommands.Count; index++)
                {
                    if (_loadArrayCommands[index].CommandTextReturn() == s.Key)
                    {
                        Command = $"[{index}][{Convert.ToInt32(number) - 1}]";
                        //_screenDelineation.ApplyCommand(index, Convert.ToInt32(number) - 1);

                        ////Process.Start((string)number);

                        ////Process.Start("http://google.com");

                        //System.Windows.Forms.WebBrowser webBrowser = new WebBrowser();

                        //webBrowser.Navigate("http://google.com");
                        //webBrowser.ScriptErrorsSuppressed = true;

                        break;
                    }
                }

                GC.Collect();
            }
        }
    }
}
