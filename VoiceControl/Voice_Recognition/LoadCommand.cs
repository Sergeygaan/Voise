﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace VoiceControlRecognition
{
    /// <summary>
    /// Класс для загрузки команд из файла
    /// </summary>
    public class LoadArrayCommands
    {
        private string _commandText;
        private List<Tuple<string, string>> _semanticResult = new List<Tuple<string, string>>();

        public LoadArrayCommands(string commandText)
        {
            _commandText = commandText;
        }

        public void AddSemantic(string command, string number)
        {
            _semanticResult.Add(new Tuple<string, string>(command, number));
        }

        public string CommandTextReturn()
        {
            return _commandText;
        }

        public List<Tuple<string, string>> SemanticResultReturn()
        {
            return _semanticResult;
        }
    }

    public class LoadCommand
    {
        private string _fileName = @"Command.xml";
        public List<LoadArrayCommands> _arrayCommands = new List<LoadArrayCommands>();

        public List<LoadArrayCommands> OpenRead()
        {
            XmlNode attr = null;

            if (!File.Exists(_fileName))
            {
                using (StreamWriter textFile = new StreamWriter(_fileName, false, Encoding.UTF8))
                {
                    textFile.WriteLine(ProjectSettings());
                }
            }

            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(_fileName);

            // получим корневой элемент
            XmlElement xRoot = xDoc.DocumentElement;

            // обход всех узлов в корневом элементе
            foreach (XmlNode xnode in xRoot)
            {
                // получаем атрибут name
                if (xnode.Attributes.Count > 0)
                {
                    attr = xnode.Attributes.GetNamedItem("name");
                    _arrayCommands.Add(new LoadArrayCommands(attr.Value));
                }

                // обходим все дочерние узлы элемента user
                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    // если узел - company
                    if (childnode.Name == "parameter")
                    {
                        Parser(childnode.InnerText);
                    }
                }
            }

            return _arrayCommands;
        }

        private void Parser(string textFromFile)
        {
            var currentText = textFromFile.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            _arrayCommands[_arrayCommands.Count - 1].AddSemantic(currentText[0], currentText[1]);
        }

        //удаление символов из строки
        private string RemoveSpaces(string inpurString, params string[] valueDeletion)
        {
            foreach (var currentValuein in valueDeletion)
            {
                inpurString = inpurString.Replace(currentValuein, string.Empty);
                inpurString = inpurString.Trim().Replace(currentValuein, string.Empty);
            }

            return inpurString;
        }

        #region Стартовые настройки

        private string ProjectSettings()
        {
            string projectSettings = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" +
            "<commands>\r\n" +
                 "<command name=\"левой\">\r\n" +
                     "   <parameter>один, 1</parameter>\r\n" +
                     "   <parameter>два, 2</parameter>\r\n" +
                     "   <parameter>три, 3</parameter>\r\n" +
                     "   <parameter>четыре, 4</parameter>\r\n" +
                     "   <parameter>пять, 5</parameter>\r\n" +
                     "   <parameter>шесть, 6</parameter>\r\n" +
                     "   <parameter>семь, 7</parameter>\r\n" +
                     "   <parameter>восемь, 8</parameter>\r\n" +
                     "   <parameter>девять, 9</parameter>\r\n" +
                 "</command>\r\n" +
                 "<command name=\"правой\">\r\n" +
                     "   <parameter>один, 1</parameter>\r\n" +
                     "   <parameter>два, 2</parameter>\r\n" +
                     "   <parameter>три, 3</parameter>\r\n" +
                     "   <parameter>четыре, 4</parameter>\r\n" +
                     "   <parameter>пять, 5</parameter>\r\n" +
                     "   <parameter>шесть, 6</parameter>\r\n" +
                     "   <parameter>семь, 7</parameter>\r\n" +
                     "   <parameter>восемь, 8</parameter>\r\n" +
                     "   <parameter>девять, 9</parameter>\r\n" +
                 "</command>\r\n" +
                 "<command name=\"двойной\">\r\n" +
                     "   <parameter>один, 1</parameter>\r\n" +
                     "   <parameter>два, 2</parameter>\r\n" +
                     "   <parameter>три, 3</parameter>\r\n" +
                     "   <parameter>четыре, 4</parameter>\r\n" +
                     "   <parameter>пять, 5</parameter>\r\n" +
                     "   <parameter>шесть, 6</parameter>\r\n" +
                     "   <parameter>семь, 7</parameter>\r\n" +
                     "   <parameter>восемь, 8</parameter>\r\n" +
                     "   <parameter>девять, 9</parameter>\r\n" +
                 "</command>\r\n" +
                 "<command name=\"сектор\">\r\n" +
                     "   <parameter>один, 1</parameter>\r\n" +
                     "   <parameter>два, 2</parameter>\r\n" +
                     "   <parameter>три, 3</parameter>\r\n" +
                     "   <parameter>четыре, 4</parameter>\r\n" +
                     "   <parameter>пять, 5</parameter>\r\n" +
                     "   <parameter>шесть, 6</parameter>\r\n" +
                     "   <parameter>семь, 7</parameter>\r\n" +
                     "   <parameter>восемь, 8</parameter>\r\n" +
                     "   <parameter>девять, 9</parameter>\r\n" +
                 "</command>\r\n" +
                 "<command name=\"масштаб\">\r\n" +
                     "   <parameter>один, 1</parameter>\r\n" +
                     "   <parameter>два, 2</parameter>\r\n" +
                     "   <parameter>три, 3</parameter>\r\n" +
                     "   <parameter>четыре, 4</parameter>\r\n" +
                     "   <parameter>пять, 5</parameter>\r\n" +
                     "   <parameter>шесть, 6</parameter>\r\n" +
                     "   <parameter>семь, 7</parameter>\r\n" +
                     "   <parameter>восемь, 8</parameter>\r\n" +
                     "   <parameter>девять, 9</parameter>\r\n" +
                 "</command>\r\n" +
                 "<command name=\"нажать\">\r\n" +
                     "   <parameter>один, 1</parameter>\r\n" +
                     "   <parameter>два, 2</parameter>\r\n" +
                     "   <parameter>три, 3</parameter>\r\n" +
                     "   <parameter>четыре, 4</parameter>\r\n" +
                     "   <parameter>пять, 5</parameter>\r\n" +
                     "   <parameter>шесть, 6</parameter>\r\n" +
                     "   <parameter>семь, 7</parameter>\r\n" +
                     "   <parameter>восемь, 8</parameter>\r\n" +
                     "   <parameter>девять, 9</parameter>\r\n" +
                 "</command>\r\n" +
                 "<command name=\"отпустить\">\r\n" +
                     "   <parameter>один, 1</parameter>\r\n" +
                     "   <parameter>два, 2</parameter>\r\n" +
                     "   <parameter>три, 3</parameter>\r\n" +
                     "   <parameter>четыре, 4</parameter>\r\n" +
                     "   <parameter>пять, 5</parameter>\r\n" +
                     "   <parameter>шесть, 6</parameter>\r\n" +
                     "   <parameter>семь, 7</parameter>\r\n" +
                     "   <parameter>восемь, 8</parameter>\r\n" +
                     "   <parameter>девять, 9</parameter>\r\n" +
                 "</command>\r\n" +
                 "<command name=\"средняя\">\r\n" +
                     "   <parameter>вверх, 1</parameter>\r\n" +
                     "   <parameter>вниз, 2</parameter>\r\n" +
                 "</command>\r\n" +
                 "<command name=\"отменить\">\r\n" +
                     "   <parameter>сектор, 1</parameter>\r\n" +
                     "   <parameter>масштаб, 2</parameter>\r\n" +
                     "   <parameter>все, 3</parameter>\r\n" +
                 "</command>\r\n" +
             "</commands>";

            return projectSettings;
        }

        #endregion
    }
}
