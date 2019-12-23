using System;
using System.ComponentModel;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using VoiceControlCommon;
using VoiceControlRecognition;
using WorkingScreen;

namespace VoiceControlCore
{
    public class ControlCore
    {
        private ListView _listView;
        private ScreenDelineation _screenDelineation;

        public ControlCore(ListView listView)
        {
            _listView = listView;

            _screenDelineation = new ScreenDelineation();
            _screenDelineation.Show();

            Recognition voiceRecognition = new Recognition("ru-RU");
            Recognition.PropertyTextLog += new PropertyChangedEventHandler(AppendLine);
            Recognition.PropertyCommand += new PropertyChangedEventHandler(Command);
        }

        /// <summary>
        /// Получение команды на обработку
        /// </summary>
        private void Command(object sender, PropertyChangedEventArgs even)
        {
            var command = ParsingText.Parsing(even.PropertyName);
            _screenDelineation.ApplyCommand(Convert.ToInt32(command[0]), Convert.ToInt32(command[1]));
        }

        /// <summary>
        /// Вывод сообщения в лог
        /// </summary>
        /// <param name="text"></param>
        /// <param name="color"></param>
        private void AppendLine(object sender, PropertyChangedEventArgs even)
        {
            var textLog = ParsingText.Parsing(even.PropertyName);

            var listViewItem = new ListViewItem
            {
                Text = textLog[0],
                BackColor = Color.FromName(textLog[1])
            };

            _listView.Items.Add(listViewItem);

            if (_listView.Items.Count > 2)
            {
                _listView.Items[_listView.Items.Count - 2].Focused = false;
                _listView.Items[_listView.Items.Count - 2].Focused = false;

                _listView.Items[_listView.Items.Count - 1].Focused = true;
                _listView.Items[_listView.Items.Count - 1].Focused = true;
                _listView.Items[_listView.Items.Count - 1].EnsureVisible();
            }

            if (_listView.Items.Count > 100)
            {
                _listView.Items.Clear();
            }

            textLog = null;
            listViewItem = null;
        }
    }
}
