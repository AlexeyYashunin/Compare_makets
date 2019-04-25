using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;

namespace Compare_makets {
    public enum file_type { WORD=0 };
    public class TFile {
        public Form1 parentForm;//указатель на родительскую форму
        public bool isExists = false;

        public List<Microsoft.Office.Interop.Word.Range> TablesRanges;
        public Microsoft.Office.Interop.Word.Application wordApp;

        public Document doc;

        public Document outFirstDoc=null;
        public Document outSecondDoc=null;

        public file_type fType;
        public string filePath;
        public string outFilePath;

        private object lockFirstFile = new object();
        private object lockSecondFile = new object();

        //public List<List<Tuple<string,List<string>>>> tables=new List<List<Tuple<string, List<string>>>>();
        public List<Tuple<string, string>> tableData = new List<Tuple<string, string>>();
        
        public int getRowsSum() {
            int prBarSize = 0;

            foreach(Microsoft.Office.Interop.Word.Table table in wordApp.Application.ActiveDocument.Tables) {
                prBarSize += table.Rows.Count;
            }

            return prBarSize;
        }

        public void Start() {
            switch(fType) {
                case file_type.WORD: {
                        TablesRanges = new List<Microsoft.Office.Interop.Word.Range>();
                        //wordApp = new Microsoft.Office.Interop.Word.Application();

                        try {
                            object misDocArg = Type.Missing;
                            object fPath = filePath;

                            doc = wordApp.Documents.OpenNoRepairDialog(ref fPath, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg);
                            //doc = wordApp.Documents.Open(ref fPath, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg);
                            //throw new IndexOutOfRangeException();
                        } catch(Exception ex) {
                            wordApp.Quit();
                            //MessageBox.Show("Error", ex.Message);
                            return;
                        }

                        bool endLine = false;
                        
                        //..........................
                        int tableIndex = 0; ;
                        int amountOfColumns = 0;

                        int columnsOffset = 0;

                        foreach(Microsoft.Office.Interop.Word.Table table in wordApp.Application.ActiveDocument.Tables) {
                            //foreach(Microsoft.Office.Interop.Word.Row tableRow in wordApp.Application.ActiveDocument.Tables[1].Rows) {

                            int rowCount = table.Rows.Count;
                            int colCount = table.Columns.Count;

                            for(int i = 1; i <= rowCount; i++) {
                                for(int j = 1; j <= colCount; j++) {
                                    Thread.Sleep(250);

                                    string tempStr = "";

                                    try {
                                        tempStr = table.Cell(i, j).Range.Text.Replace("\r", "").Replace("\a", "").Replace("\r\a", "").Replace("\v", "");
                                    } catch(System.Exception ex) {
                                    }

                                    if(i == 1) {
                                        tableData.Add(new Tuple<string, string>(tempStr, ""));
                                    } else if(tempStr != tableData[j + columnsOffset - 1].Item1) {
                                        if(i == 2) {
                                            tableData[j + columnsOffset - 1] = new Tuple<string, string>(tableData[j + columnsOffset - 1].Item1, tempStr);
                                        } else {
                                            try {
                                                tempStr = table.Cell(i, j).Range.Text.Replace("\r", ",").Replace("\a", ",").Replace("\r\a", ",").Replace("\v", ",");
                                            } catch(System.Exception ex) {
                                            }

                                            tableData[j + columnsOffset - 1] = new Tuple<string, string>(tableData[j + columnsOffset - 1].Item1, tableData[j + columnsOffset - 1].Item2 + "," + tempStr);
                                        }
                                    }
                                }
                            }

                            columnsOffset += table.Rows[1].Cells.Count;

                            tableIndex++;
                        }

                        /*if(doc != null) {
                            doc.Close();
                        }
                        wordApp.Quit();*/


                        break;
                    }
            }
        }

        public TFile(Form1 parentForm, string filePath, string outFileName, file_type fType) {
            if(System.IO.File.Exists(filePath)==false) {
                isExists = false;
                return;
            } else {
                isExists = true;
            }

            this.parentForm = parentForm;

            this.fType = fType;
            this.filePath = filePath;

            //outFilePath = System.IO.Directory.GetCurrentDirectory() + "\\" + "[OUT]Data" + "\\"+ System.IO.Path.GetFileName(filePath);
            outFilePath = System.IO.Directory.GetCurrentDirectory() + "\\" + "[OUT]Data" + "\\" + outFileName + System.IO.Path.GetExtension(filePath);

            string outDataFolder = System.IO.Directory.GetCurrentDirectory() + "\\" + "[OUT]Data";

            if (System.IO.Directory.Exists(outDataFolder)==false) {
                //создаем путь для выходных файлов
                System.IO.Directory.CreateDirectory(outDataFolder);

                //даем права на запись в папку
                System.IO.File.SetAttributes(outDataFolder, System.IO.FileAttributes.Normal);
            }

            try{
                System.IO.File.Copy(filePath, outFilePath, true);//делаем дубликат файла для последующего использования
            } catch (System.Exception ex){          	
            }

            switch(fType) {
                case file_type.WORD: {
                        wordApp = new Microsoft.Office.Interop.Word.Application();
                        //wordApp.Options.DefaultHighlightColorIndex= Microsoft.Office.Interop.Word.WdColorIndex.wdGreen;

                        break;
                    }
            }
        }

        public void Compare(TFile tFile) {
            parentForm.backgroundWorker1.ReportProgress(0);

            List<string> lostWords = new List<string>();
            
            List<Tuple<string, string[]>> tdFirstFileClear = new List<Tuple<string, string[]>>();
            List<Tuple<string, string[]>> tdSecondFileClear = new List<Tuple<string, string[]>>();

            List<string> codeTNVED = new List<string>();

            //Document docTemp = wordApp.Documents.OpenNoRepairDialog(FileName: tFile.filePath, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false, NoEncodingDialog: true);

            //System.IO.File.SetAttributes(tFile.outFilePath, System.IO.FileAttributes.Normal);

            //Document docTemp = wordApp.Documents.OpenNoRepairDialog(FileName: tFile.outFilePath, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false, NoEncodingDialog: true);

            object misDocArg = Type.Missing;
            object fFirstFilePath = outFilePath;
            object fSecondFilePath = tFile.outFilePath;

            /*Document outFirstDoc = wordApp.Documents.OpenNoRepairDialog(FileName: tFile.outFilePath);
            Document outSecondDoc = wordApp.Documents.OpenNoRepairDialog(FileName: outFilePath);*/

            outSecondDoc = wordApp.Documents.OpenNoRepairDialog(fSecondFilePath, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg);
            outFirstDoc = wordApp.Documents.OpenNoRepairDialog(fFirstFilePath, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg, ref misDocArg);
            
            //wordApp.Options.DefaultHighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdGreen;

            //1й файл
            Microsoft.Office.Interop.Word.Range rangeFFile = outFirstDoc.Content;
            Microsoft.Office.Interop.Word.Find findFFile = rangeFFile.Find;

            //выделяем все слова черным цветом в 1м файле
            rangeFFile.Select();
            rangeFFile.Font.Color = WdColor.wdColorBlack;

            //распарсиваем текст 1го файла на слова, сохраняем и выделяем КОДЫ ТН ВЭД в 1м файле
            for(int i=0;i<tableData.Count;i++){
                System.Windows.Forms.Application.DoEvents();

                string[] splitValues = tableData[i].Item2.Split(',',':',';');

                List<string> sVClear = new List<string>();

                for(int k = 0; k < splitValues.Length; k++) {
                    if(string.IsNullOrWhiteSpace(splitValues[k])==false) {
                        if(i==0){
                            codeTNVED.Add(splitValues[k]);
                            splitValues[k] = splitValues[k].Replace(" ", string.Empty);

                            //сразу выделяем зеленым цветом КОД ТН ВЭД в 1м файле
                            //findFFile.Text = codeTNVED[codeTNVED.Count-1];
                            findFFile.Replacement.Font.Color = WdColor.wdColorGreen;

                            object findText = codeTNVED[codeTNVED.Count - 1];
                            object missing = System.Reflection.Missing.Value;
                            object replace = WdReplace.wdReplaceAll;//заменяем все вхождения
                            object replaceWith = codeTNVED[codeTNVED.Count - 1];

                            findFFile.Execute(
                                ref findText, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                                ref replace, ref missing, ref missing, ref missing, ref missing
                            );

                            rangeFFile = outFirstDoc.Content;
                            findFFile = rangeFFile.Find;
                        }

                        sVClear.Add(splitValues[k]);
                    }
                }

                Tuple<string, string[]> tempTuple = new Tuple<string, string[]>(tableData[i].Item1, sVClear.ToArray());
                tdFirstFileClear.Add(tempTuple);

                //outFirstDoc.Save();
            }

            //сохраняем 1й файл
            outFirstDoc.Save();

            /*...............................
            Действия в блоке:
            1. Проверяем все слова, чтобы не превышали 250 символов
            ...............................*/

            //в начале столбцы
            for(int i = 0; i < tdFirstFileClear.Count; i++) {
                int wordLength = tdFirstFileClear[i].Item1.Length;

                if(wordLength > 250) {
                    do {
                        wordLength = wordLength - (wordLength/2);//получаем макс. часть длины слова при делении на 2                        
                    } while(wordLength > 250);

                    //делим слова
                    int partStrCount = (tdFirstFileClear[i].Item1.Length / wordLength);
                    int ostWordLength = (tdFirstFileClear[i].Item1.Length - (partStrCount* wordLength));

                    List<string> tempList = new List<string>(tdFirstFileClear[i].Item2.ToArray());

                    for(int j=1;j<partStrCount;j++){//берем со 2-го слова и добавляем в конец item2, потом перезапишем item1 после цикла
                        string subStr = tdFirstFileClear[i].Item1.Substring(j * wordLength, wordLength);

                        tempList.Add(subStr);
                    }

                    //добавляем остаточное слово, если требуется
                    if(ostWordLength > 0) {
                        string subStr = tdFirstFileClear[i].Item1.Substring(tdFirstFileClear[i].Item1.Length - ostWordLength - 1, ostWordLength);

                        tempList.Add(subStr);
                    }

                    //пересоздаем элемент
                    Tuple<string, string[]> tempTuple = new Tuple<string, string[]>(tdFirstFileClear[i].Item1.Substring(0, wordLength), tempList.ToArray());
                    tdFirstFileClear[i] = tempTuple;

                    /*tdFirstFileClear.RemoveAt(i);
                    tdFirstFileClear.Insert(i,tempTuple);*/
                }
            }

            //теперь значения таблицы
            for(int i = 0; i < tdFirstFileClear.Count; i++) {
                for(int j = 0; j < tdFirstFileClear[i].Item2.Length; j++) {
                    int wordLength = tdFirstFileClear[i].Item2[j].Length;

                    if(wordLength > 250) {
                        do {
                            wordLength = wordLength - (wordLength / 2);//получаем макс. часть длины слова при делении на 2                        
                        } while(wordLength > 250);

                        //делим слова
                        int partStrCount = (tdFirstFileClear[i].Item2[j].Length / wordLength);
                        int ostWordLength = (tdFirstFileClear[i].Item2[j].Length - (partStrCount * wordLength));

                        List<string> tempList = new List<string>(tdFirstFileClear[i].Item2.ToArray());
                        tempList.RemoveAt(j);//удаляем строку > 250 символов

                        for(int k = 0; k < partStrCount; k++) {
                            string subStr = tdFirstFileClear[i].Item2[j].Substring(k * wordLength, wordLength);

                            tempList.Add(subStr);
                        }

                        //добавляем остаточное слово, если требуется
                        if(ostWordLength>0){
                            string subStr = tdFirstFileClear[i].Item2[j].Substring(tdFirstFileClear[i].Item2[j].Length - ostWordLength - 1, ostWordLength);

                            tempList.Add(subStr);
                        }

                        //пересоздаем элемент
                        Tuple<string, string[]> tempTuple = new Tuple<string, string[]>(tdFirstFileClear[i].Item1, tempList.ToArray());
                        tdFirstFileClear[i] = tempTuple;

                        /*tdFirstFileClear.RemoveAt(i);
                        tdFirstFileClear.Insert(i,tempTuple);*/
                    }
                }
            }

            ///////////////////////////////////////
            ///////////////////////////////////////

            /*...............................
            Действия в блоке:
            1. Ищем табличные слова 1го файла во 2м с выделением зеленым цветом
            2. Ищем табличные слова 1го файла в 1м с выделением зеленым цветом(если нашлось во 2м) и красным цветом(если не нашлось во 2м)
            ...............................*/
            bool isFind = false;

            //2й файл
            Microsoft.Office.Interop.Word.Range rangeSFile = outSecondDoc.Content;
            Microsoft.Office.Interop.Word.Find findSFile = rangeSFile.Find;

            //выделяем все слова черным цветом во 2м файле
            rangeSFile.Select();
            rangeSFile.Font.Color = WdColor.wdColorBlack;

            //подготавливаем 1й файл
            rangeFFile = outFirstDoc.Content;
            findFFile = rangeFFile.Find;

            int pBarMaxValue = 0;

            //считаем числовой предел progressBara
            for (int i = 0; i < tdFirstFileClear.Count; i++) {
                pBarMaxValue += tdFirstFileClear[i].Item2.Length;
            }

            int curProgressPos = 0;//текущая позиция progressBara

            for(int i = 0; i < tdFirstFileClear.Count; i++) {
                //ищем названия столбцов 1го файла во 2м
                isFind = false;

                object findText = tdFirstFileClear[i].Item1;//лучше через ref к findText, чем присвоением к полю find.Text, т.к. потом текста не всегда находится

                //findSFile.Text = tdFirstFileClear[i].Item1;
                findSFile.Replacement.Font.Color = WdColor.wdColorGreen;

                object missing = System.Reflection.Missing.Value;
                object replace = WdReplace.wdReplaceAll;//заменяем все вхождения
                object replaceWith = tdFirstFileClear[i].Item1;

                isFind = findSFile.Execute(
                    ref findText, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                    ref replace, ref missing, ref missing, ref missing, ref missing
                );

                //выделяем слово в 1м файле
                //findFFile.Text = tdFirstFileClear[i].Item1;
                findText = tdFirstFileClear[i].Item1;
                findFFile.Replacement.Font.Color = WdColor.wdColorGreen;
                replaceWith = tdFirstFileClear[i].Item1;

                findFFile.Execute(
                    ref findText, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                    ref replace, ref missing, ref missing, ref missing, ref missing
                );

                rangeFFile = outFirstDoc.Content;
                findFFile = rangeFFile.Find;

                //если название столбца 1го файла не найдено во 2м, то пробуем без пробелов
                if (isFind == false) {
                    //string noSpaceStr = tdFirstFileClear[i].Item1.Replace(" ", string.Empty);
                    object noSpaceStr = Regex.Replace(tdFirstFileClear[i].Item1, @"[^\w-]", "");//удаляем все символы кроме букв, цифр и '-'

                    //findSFile.Text = noSpaceStr;

                    rangeSFile = outSecondDoc.Content;
                    findSFile = rangeSFile.Find;
                    replaceWith = noSpaceStr;

                    isFind = findSFile.Execute(
                        ref noSpaceStr, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                        ref replace, ref missing, ref missing, ref missing, ref missing
                    );

                    if (isFind == false) {
                        lostWords.Add(tdFirstFileClear[i].Item1);

                        //сразу выделяем красным цветом не найденное слово 1го файла в 1м файле
                        //findFFile.Text = tdFirstFileClear[i].Item1;
                        findText = tdFirstFileClear[i].Item1;
                        findFFile.Replacement.Font.Color = WdColor.wdColorRed;
                        replaceWith = tdFirstFileClear[i].Item1;

                        findFFile.Execute(
                            ref findText, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                            ref replace, ref missing, ref missing, ref missing, ref missing
                        );

                        rangeFFile = outFirstDoc.Content;
                        findFFile = rangeFFile.Find;
                    }
                }

                rangeSFile = outSecondDoc.Content;
                findSFile = rangeSFile.Find;

                //ищем табличные слова 1го файла во 2м
                for(int j = 0; j < tdFirstFileClear[i].Item2.Length; j++) {
                    isFind = false;                   

                    /*if (tdFirstFileClear[i].Item2[j] == " 433277") {
                        string temp = "";
                    }*/

                    //findSFile.Text = tdFirstFileClear[i].Item2[j];
                    findText = tdFirstFileClear[i].Item2[j];
                    findSFile.Replacement.Font.Color = WdColor.wdColorGreen;
                    replaceWith = tdFirstFileClear[i].Item2[j];

                    isFind = findSFile.Execute(
                        ref findText, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                        ref replace, ref missing, ref missing, ref missing, ref missing
                    );

                    //выделяем табличное слово в 1м файле
                    //findFFile.Text = tdFirstFileClear[i].Item2[j];
                    findText = tdFirstFileClear[i].Item2[j];
                    findFFile.Replacement.Font.Color = WdColor.wdColorGreen;
                    replaceWith = tdFirstFileClear[i].Item2[j];

                    findFFile.Execute(
                        ref findText, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                        ref replace, ref missing, ref missing, ref missing, ref missing
                    );

                    rangeFFile = outFirstDoc.Content;
                    findFFile = rangeFFile.Find;

                    rangeSFile = outSecondDoc.Content;
                    findSFile = rangeSFile.Find;

                    //если табличное слово 1го файла не найдено во 2м, то пробуем без пробелов
                    if(isFind == false) {
                        //string noSpaceStr = tdFirstFileClear[i].Item2[j].Replace(" ", string.Empty);
                        //string noSpaceStr = Regex.Replace(tdFirstFileClear[i].Item2[j], "[^A-Za-z0-9-]", "");//удаляем все символы кроме букв, цифр и '-'
                        object noSpaceStr = Regex.Replace(tdFirstFileClear[i].Item2[j], @"[^\w-]", "");//удаляем все символы кроме букв, цифр и '-'

                        //findSFile.Text = noSpaceStr;

                        rangeSFile = outSecondDoc.Content;
                        findSFile = rangeSFile.Find;

                        findSFile.Replacement.Font.Color = WdColor.wdColorGreen;

                        replaceWith = noSpaceStr;

                        isFind = findSFile.Execute(
                            ref noSpaceStr, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                            ref replace, ref missing, ref missing, ref missing, ref missing
                        );

                        if(isFind == false) {
                            lostWords.Add(tdFirstFileClear[i].Item2[j]);

                            //сразу выделяем красным цветом не найденное слово 1го файла в 1м файле
                            //findFFile.Text = tdFirstFileClear[i].Item2[j];
                            findText = tdFirstFileClear[i].Item2[j];                            
                            findFFile.Replacement.Font.Color = WdColor.wdColorRed;
                            replaceWith = tdFirstFileClear[i].Item2[j];

                            findFFile.Execute(
                                ref findText, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                                ref replace, ref missing, ref missing, ref missing, ref missing
                            );

                            rangeFFile = outFirstDoc.Content;
                            findFFile = rangeFFile.Find;
                        }
                    }

                    rangeSFile = outSecondDoc.Content;
                    findSFile = rangeSFile.Find;

                    rangeFFile = outFirstDoc.Content;
                    findFFile = rangeFFile.Find;

                    //достигаем 99%, а оставшийся 1% добавляем на выходе из функции
                    if((curProgressPos+1)!=pBarMaxValue){
                        curProgressPos += 1;
                    }

                    float percent = (curProgressPos * 100) / pBarMaxValue;

                    /*if(percent==12) {
                        string te = "";
                    }*/

                    if(percent % 10 == 0) {
                        outFirstDoc.Save();
                        outSecondDoc.Save();
                    }

                    parentForm.backgroundWorker1.ReportProgress((int)percent);
                }
            }

            outFirstDoc.Save();
            outSecondDoc.Save();

            try{
                if(outFirstDoc != null) {
                    outFirstDoc.Close();
                }
            } catch (System.Exception ex){            	
            }
            
            try{
                if (outSecondDoc != null) {
                    outSecondDoc.Close();
                }
            }catch (System.Exception ex){                
            }            
            
            System.IO.File.WriteAllLines(System.IO.Directory.GetCurrentDirectory() + "\\" + "[OUT]Data" + "\\" + "Lost_words.txt", lostWords.ToArray());

            parentForm.backgroundWorker1.ReportProgress(100);
        }

        public void Destroy() {
            try {
                if (doc != null) {                    
                    doc.Close();
                }
            } catch (System.Exception ex) {
            }

            try {
                if(outFirstDoc != null) {
                    outFirstDoc.Close();
                }
            } catch(System.Exception ex) {
            }

            try {
                if(outSecondDoc != null) {
                    outSecondDoc.Close();
                }
            } catch(System.Exception ex) {
            }

            try {
                if(wordApp != null) {
                    wordApp.Quit();
                }
            } catch (System.Exception ex){            	
            }           
        }
    }
}
