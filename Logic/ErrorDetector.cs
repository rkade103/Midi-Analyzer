using OfficeOpenXml;
using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;

namespace Midi_Analyzer.Logic
{
    class ErrorDetector
    {
        private FileInfo excerptFileInfo;
        private ExcelPackage excerptWb;

        private readonly int FROZEN_ROWS = 10;
        private readonly string[] headers = { "Line number", "Note", "Duration", "Include? (Y/N)",
            "Include TL", "Include Dyn.", "Include Art.", "Include N.D.", "Space for barline", "Graph Width",
            "Vel. Graph Width", "X-axis limit"}; //The last one may be removed, depending on the sheet reader.
        private readonly int LINE_NUMBER_COL = 1;
        private readonly int NOTE_COL = 2;
        private readonly int DURATION_COL = 3;
        private readonly int INCLUDE_COL = 4;
        private readonly int INCLUDE_TL_COL = 5;
        private readonly int INCLUDE_DYN_COL = 6;
        private readonly int INCLUDE_ART_COL = 7;
        private readonly int INCLUDE_ND_COL = 8;
        private readonly int SPACE_BARLINE_COL = 9;
        private readonly int GRAPH_WIDTH_COL = 10;
        private readonly int VEL_GRAPH_WIDTH_COL = 11;
        private readonly int X_AXIS_LIMIT_COL = 12;

        private readonly string NO_ERROR_STRING = ""; //This string represents if no errors were detected.

        public ErrorDetector(string path)
        {
            excerptFileInfo = new FileInfo(path);
            excerptWb = new ExcelPackage(excerptFileInfo);
        }

        public ErrorDetector()
        {
            
        }

        public string CheckExcerptSheetForErrors()
        {
            string message = NO_ERROR_STRING;
            ExcelWorksheet sheet = excerptWb.Workbook.Worksheets[1];

            message = CheckExcerptSheetLineNumberValues(sheet);
            if (message != NO_ERROR_STRING) //Here, you should pass the excerpt sheet worksheet instead of the path.
            {
                return message;
            }
            message = CheckExcerptSheetNoteCol(sheet);
            if(message != NO_ERROR_STRING)
            {
                return message;
            }
            message = CheckExcerptSheetDurationCol(sheet);
            if(message != NO_ERROR_STRING)
            {
                return message;
            }
            message = CheckExcerptSheetIncludeCols(sheet);
            if(message != NO_ERROR_STRING)
            {
                return message;
            }
            message = CheckExcerptSheetLastTLArtRow(sheet);
            if(message != NO_ERROR_STRING)
            {
                return message;
            }
            message = CheckExcerptSheetSpacingValues(sheet);
            if(message != NO_ERROR_STRING)
            {
                return message;
            }
            message = CheckExcerptSheetGraphWidthValues(sheet);
            if(message != NO_ERROR_STRING)
            {
                return message;
            }
            excerptWb.Save();
            return message;
        }

        /// <summary>
        /// Checks if the given excerpt sheet follows the conventions and assumptions of the analyzer.
        /// </summary>
        /// <param name="path">A path to the excerpt sheet.</param>
        /// <returns>A list of the incorrect headers.</returns>
        public List<string> CheckExcerptSheetStructure()
        {
            ExcelWorksheet sheet = excerptWb.Workbook.Worksheets[1];
            int i = 0;
            List<string> badHeaders = new List<string>();
            while(i < headers.Length)
            {
                if(!(sheet.Cells[1, i+1].Text.Trim() == headers[i]))
                {
                    badHeaders.Add(sheet.Cells[1, i+1].Text);
                }
                i++;
            }
            return badHeaders;
        }

        /// <summary>
        /// Checks if all line numbers have a corresponding note.
        /// </summary>
        /// <param name="path">Path to the excerpt sheet.</param>
        /// <returns>True represents that all numbers have a note (no errors), and false the opposite.</returns>
        public bool CheckExcerptSheetForEmptyValues()
        {
            ExcelWorksheet sheet = excerptWb.Workbook.Worksheets[1];
            int i = 2;
            int rowCount = sheet.Dimension.End.Row;
            if(sheet.Cells[i, GRAPH_WIDTH_COL].Text.ToLower().Trim() == "" || sheet.Cells[i, VEL_GRAPH_WIDTH_COL].Text.ToLower().Trim() == "" || 
                sheet.Cells[i, X_AXIS_LIMIT_COL].Text.ToLower().Trim() == "")
            {
                return false;
            }
            while (i < rowCount || sheet.Cells[i, LINE_NUMBER_COL].Text.ToLower().Trim() != "end")
            {
                for(int j=1; j <= headers.Length - 3; j++)
                {
                    if(sheet.Cells[i, j].Text.Trim() == "")
                    {
                        return false;
                    }
                }
                i++;
            }
            return true;
        }

        /// <summary>
        /// Check the line number column for invalid values.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public string CheckExcerptSheetLineNumberValues(ExcelWorksheet sheet)
        {
            int i = 2;
            int rowCount = sheet.Dimension.End.Row;
            string message = NO_ERROR_STRING;
            string currentIndexValue = sheet.Cells[i, LINE_NUMBER_COL].Text.Trim().ToLower();
            while (i < rowCount && currentIndexValue != "end")
            {
                if (!IsDigitsOnly(sheet.Cells[i, LINE_NUMBER_COL].Text.Trim()))
                {
                    message = "There was an invalid value detected in the Line Number column, at row " + i + ".\nPlease only provide a positive whole " +
                        "number as the values for the line numbers. \nOnce the last line number has been placed, input the word END in the cell directly" +
                        "below it.";
                    return message;
                }
                i++;
                currentIndexValue = sheet.Cells[i, LINE_NUMBER_COL].Text.Trim().ToLower();
            }
            return message;
        }

        /// <summary>
        /// Checks that every line number has a note value. Note that it does not check the validity of the note value. 
        /// </summary>
        /// <param name="sheet">The excerpt worksheet.</param>
        /// <returns>A string. An empty string represents no errors, whereas a custom one represents an error.</returns>
        public string CheckExcerptSheetNoteCol(ExcelWorksheet sheet)
        {
            int i = 2;
            int rowCount = sheet.Dimension.End.Row;
            string message = NO_ERROR_STRING;
            while (i < rowCount || sheet.Cells[i, LINE_NUMBER_COL].Text.ToLower().Trim() != "end")
            {
                if(!CheckNote(sheet.Cells[i, NOTE_COL].Text.Trim())){
                    message = "There was an invalid value detected in the Note column, at row " + i + ". Please make sure the notes have the " +
                        "following structure:\n" +
                        "Letter - Number - # (optional sharp)\n" +
                        "Please make sure a note value is present for every number in the line number column.";
                    return message;
                }
                i++;
            }
            return message;
        }

        /// <summary>
        /// Cheks that the entries in the duration column are valid (checks that it is a number).
        /// </summary>
        /// <param name="sheet">The excerpt worksheet.</param>
        /// <returns>A string. An empty string represents no errors, wheras a custome one represents an error.</returns>
        public string CheckExcerptSheetDurationCol(ExcelWorksheet sheet)
        {
            int i = 2;
            int rowCount = sheet.Dimension.End.Row;
            string message = NO_ERROR_STRING;
            double duration;
            string stringDuration;
            while (i < rowCount || sheet.Cells[i, LINE_NUMBER_COL].Text.ToLower().Trim() != "end")
            {
                stringDuration = sheet.Cells[i, DURATION_COL].Text.Trim();
                if (!Double.TryParse(stringDuration, out duration))
                {
                    if (stringDuration.Contains("-"))
                    {
                        message = "There was an invalid value detected in the duration column, at row " + i + ".\nPlease only provide a positive " +
                        "number as the values for the durations. These can be fractional numbers or whole numbers.";
                        return message;
                    }
                    string[] numbers = stringDuration.Split('/');
                    if(numbers.Length <= 1 || numbers.Length > 2)
                    {
                        message = "There was an invalid value detected in the duration column, at row " + i + ".\nPlease only provide a positive " +
                        "number as the values for the durations. These can be fractional numbers or whole numbers.";
                        return message;
                    }
                    else
                    {
                        int nominator;
                        int denominator;
                        if(!int.TryParse(numbers[0], out nominator) || !int.TryParse(numbers[1], out denominator))
                        {
                            message = "There was an invalid value detected in the duration column, at row " + i + ".\nPlease only provide a positive " +
                        "number as the values for the durations. These can be fractional numbers or whole numbers.";
                            return message;
                        }
                        if (denominator == 0)
                        {
                            message = "There was an invalid value detected in the duration column, at row " + i + ".\nPlease only provide a positive " +
                        "number as the values for the durations. These can be fractional numbers or whole numbers. Please make sure that the denominators " +
                        "are not 0.";
                            return message;
                        }
                    }
                }
                else
                {
                    if (duration <= 0)
                    {
                        message = "There was an invalid value detected in the duration column, at row " + i + ".\nPlease only provide a positive " +
                            "number as the values for the durations. These can be fractional numbers or whole numbers.";
                        return message;
                    }
                }
                i++;
            }
            return message;
        }

        /// <summary>
        /// Checks all of the include columns in the excerpt sheet and make sure there are only Y/N values.
        /// </summary>
        /// <param name="sheet">The excerpt sheet./param>
        /// <returns>A string message. An empty string represents no errors, whereas a custom string represents an error.</returns>
        public string CheckExcerptSheetIncludeCols(ExcelWorksheet sheet)
        {
            int i = 2;
            int rowCount = sheet.Dimension.End.Row;
            string message = NO_ERROR_STRING;
            while (i < rowCount || sheet.Cells[i, LINE_NUMBER_COL].Text.ToLower().Trim() != "end")
            {
                if (sheet.Cells[i, INCLUDE_COL].Text.Trim().ToLower() != "y" && sheet.Cells[i, INCLUDE_COL].Text.Trim().ToLower() != "n")
                {
                    message = "There was an invalid value detected in the first Include column (Column D), at row " + i + ". Please make sure there is only " +
                        "Y and N values in the column (it is not case sensitive).";
                    return message;
                }
                if (sheet.Cells[i, INCLUDE_TL_COL].Text.Trim().ToLower() != "y" && sheet.Cells[i, INCLUDE_TL_COL].Text.Trim().ToLower() != "n")
                {
                    message = "There was an invalid value detected in the second Include column (Column E), at row " + i + ". Please make sure there is only " +
                        "Y and N values in the column (it is not case sensitive).";
                    return message;
                }
                if (sheet.Cells[i, INCLUDE_DYN_COL].Text.Trim().ToLower() != "y" && sheet.Cells[i, INCLUDE_DYN_COL].Text.Trim().ToLower() != "n")
                {
                    message = "There was an invalid value detected in the third Include column (Column F), at row " + i + ". Please make sure there is only " +
                        "Y and N values in the column (it is not case sensitive).";
                    return message;
                }
                if (sheet.Cells[i, INCLUDE_ART_COL].Text.Trim().ToLower() != "y" && sheet.Cells[i, INCLUDE_ART_COL].Text.Trim().ToLower() != "n")
                {
                    message = "There was an invalid value detected in the fourth Include column (Column G), at row " + i + ". Please make sure there is only " +
                        "Y and N values in the column (it is not case sensitive).";
                    return message;
                }
                if (sheet.Cells[i, INCLUDE_ND_COL].Text.Trim().ToLower() != "y" && sheet.Cells[i, INCLUDE_ND_COL].Text.Trim().ToLower() != "n")
                {
                    message = "There was an invalid value detected in the fifth Include column (Column H), at row " + i + ". Please make sure there is only " +
                        "Y and N values in the column (it is not case sensitive).";
                    return message;
                }
                i++;
            }
            return message;
        }

        public string CheckExcerptSheetLastTLArtRow(ExcelWorksheet sheet)
        {
            int i = sheet.Dimension.End.Row;
            string message = NO_ERROR_STRING;
            while (i > 0 || sheet.Cells[i, LINE_NUMBER_COL].Text.ToLower().Trim() != headers[0])
            {
                if (IsDigitsOnly(sheet.Cells[i, LINE_NUMBER_COL].Text.Trim()))
                {
                    if(sheet.Cells[i, INCLUDE_TL_COL].Text.Trim().ToLower() != "n" || sheet.Cells[i, INCLUDE_ART_COL].Text.Trim().ToLower() != "n")
                    {
                        message = "The last values in the Include TL and Include Art column must always be N, given it those variables cannot be " +
                            "generated for the last notes. This has been automatically changed now. Please press the \"Analyze\" button again to run" +
                            "the analysis.";
                        sheet.Cells[i, INCLUDE_TL_COL].Value = "N";
                        sheet.Cells[i, INCLUDE_ART_COL].Value = "N";
                        excerptWb.Save();
                        return message;
                    }
                    else
                    {
                        break;
                    }
                }
                i--;
            }
            return message;
        }

        /// <summary>
        /// Checks the spacing column for invalid values. Valid values are only that of positive doubles.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public string CheckExcerptSheetSpacingValues(ExcelWorksheet sheet)
        {
            int i = 2;
            int rowCount = sheet.Dimension.End.Row;
            string message = NO_ERROR_STRING;
            double number;
            while (i < rowCount || sheet.Cells[i, LINE_NUMBER_COL].Text.ToLower().Trim() != "end")
            {
                if(Double.TryParse(sheet.Cells[i, SPACE_BARLINE_COL].Text.Trim(), out number))
                {
                    if(number <= 0)
                    {
                        message = "There was an invalid value detected in the Space for Barline column, at row " + i + ".\nPlease only provide a positive " +
                        "number as the values for the space for barline.";
                        return message;
                    }
                }
                else
                {
                    message = "There was an invalid value detected in the Space for Barline column, at row " + i + ".\nPlease only provide a positive " +
                        "number as the values for the space for barline.";
                    return message;
                }
                i++;
            }
            return message;
        }

        /// <summary>
        /// Check the graph width number for invalid values.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public string CheckExcerptSheetGraphWidthValues(ExcelWorksheet sheet)
        {
            if(!IsDigitsOnly(sheet.Cells[2, GRAPH_WIDTH_COL].Text.Trim()))
            {
                string message = "There was an invalid value detected in the Graph Width column (column J), at row 2.\nPlease only provide a positive whole " +
                        "number as the value for the graph width. \nNote that only the first number below the column header is considered.";
                return message;
            }
            if (!IsDigitsOnly(sheet.Cells[2, VEL_GRAPH_WIDTH_COL].Text.Trim()))
            {
                string message = "There was an invalid value detected in the Velocity Graph Width column (column K), at row 2.\nPlease only provide a positive whole " +
                        "number as the value for the velocity graph width. \nNote that only the first number below the column header is considered.";
                return message;
            }
            if (!IsDigitsOnly(sheet.Cells[2, X_AXIS_LIMIT_COL].Text.Trim()))
            {
                string message = "There was an invalid value detected in the X-axis limit column (column L), at row 2.\nPlease only provide a positive whole " +
                        "number as the value for the X-axis limit. \nNote that only the first number below the column header is considered.";
                return message;
            }
            return NO_ERROR_STRING;
        }

        /// <summary>
        /// A basic error detection algorithm. It checks if all the notes in the playthrough are correct. Should an error be detected,
        /// it restarts from the end of the sheet back towards the top. This way, as many correct notes as possible can be detected.
        /// </summary>
        /// <param name="midiWb">The workbook to scan for errors.</param>
        /// <param name="excerptWb">Th workbook containing the excerpt.</param>
        /// <returns></returns>
        public List<string> ScanWorkbookForErrors(ExcelPackage midiWb, ExcelPackage excerptWb)
        {
            List<string> badSheets = new List<string>();

            ExcelWorksheet midiSheet = null;
            ExcelWorksheet excerptSheet = excerptWb.Workbook.Worksheets[1];
            for(int i = 1; i <= midiWb.Workbook.Worksheets.Count; i++)
            {
                midiSheet = midiWb.Workbook.Worksheets[i];
                if(!DetectGoodPlaythrough(midiSheet, excerptSheet))
                {
                    badSheets.Add(midiSheet.Name);
                }
            }
            midiWb.Save();
            return badSheets;
        }

        /// <summary>
        /// Checks if the note in the midisheet match the excerpt sheet.
        /// </summary>
        /// <param name="midiSheet">The sheet of the sample, representing the playthrough.</param>
        /// <param name="excerptSheet">The excerpt sheet, representing the score.</param>
        /// <returns></returns>
        public bool DetectGoodPlaythrough(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            string header = "";
            int excerptIndex = 2;
            int midiIndex = FROZEN_ROWS + 1;
            while (header != "end_of_file")
            {
                header = midiSheet.Cells[midiIndex, 4].Text.Trim().ToLower();
                if(header == "note_on_c" && int.Parse(midiSheet.Cells[midiIndex, 8].Text.Trim()) != 0)
                {
                    if(excerptSheet.Cells[excerptIndex, LINE_NUMBER_COL].Text.Trim().ToLower() == "end" || excerptSheet.Cells[excerptIndex, NOTE_COL].Text.Trim().ToLower() == "end"){
                        excerptIndex = 2; //Resets the excerpt, in case the person has multiple attempts on the same track.
                    }
                    if (midiSheet.Cells[midiIndex, 7].Text.Trim().ToLower() == excerptSheet.Cells[excerptIndex, NOTE_COL].Text.Trim().ToLower())
                    {
                        midiSheet.Cells[midiIndex, 11].Value = excerptSheet.Cells[excerptIndex, INCLUDE_COL].Value;
                        midiSheet.Cells[midiIndex, 12].Value = excerptSheet.Cells[excerptIndex, LINE_NUMBER_COL].Value;
                        midiSheet.Cells[midiIndex, 13].Value = excerptSheet.Cells[excerptIndex, DURATION_COL].Value;
                        excerptIndex++;
                    }
                    else
                    {
                        midiSheet.Cells[midiIndex, 14].Value = "ERROR";
                        return DetectGoodPlaythroughReversed(midiSheet, excerptSheet); //error detected
                    }
                }
                else if(header == "note_on_c" && int.Parse(midiSheet.Cells[midiIndex, 8].Text.Trim()) == 0)
                {
                    Console.WriteLine("NOTE_ON VELOCITY 0 DETECTED AT: " + midiIndex);
                }
                midiIndex++;
            }
            return true; //No errors found
        }

        /// <summary>
        /// Checks if the notes in the midisheet match the excerpt sheet, in reverse order.
        /// </summary>
        /// <param name="midiSheet">The sheet of the sample, representing the playthrough.</param>
        /// <param name="excerptSheet">The excerpt sheet, representing the score.</param>
        /// <returns></returns>
        public bool DetectGoodPlaythroughReversed(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            string header = "";
            int excerptIndex = excerptSheet.Dimension.End.Row;
            int midiIndex = midiSheet.Dimension.End.Row;
            while (header != "start_track")
            {
                header = midiSheet.Cells[midiIndex, 4].Text.Trim().ToLower();
                if (header == "note_on_c" && int.Parse(midiSheet.Cells[midiIndex, 8].Text.Trim()) != 0)
                {
                    if (excerptSheet.Cells[excerptIndex, LINE_NUMBER_COL].Text.Trim().ToLower() == "end" || excerptSheet.Cells[excerptIndex, NOTE_COL].Text.Trim().ToLower() == "end")
                    {
                        //excerptIndex = excerptSheet.Dimension.End.Row; //Resets the excerpt, in case the person has multiple attempts on the same track.
                        excerptIndex--;
                    }
                    else if (midiSheet.Cells[midiIndex, 7].Text.Trim().ToLower() == excerptSheet.Cells[excerptIndex, NOTE_COL].Text.Trim().ToLower())
                    {
                        midiSheet.Cells[midiIndex, 11].Value = excerptSheet.Cells[excerptIndex, INCLUDE_COL].Value;
                        midiSheet.Cells[midiIndex, 12].Value = excerptSheet.Cells[excerptIndex, LINE_NUMBER_COL].Value;
                        midiSheet.Cells[midiIndex, 13].Value = excerptSheet.Cells[excerptIndex, DURATION_COL].Value;
                        excerptIndex--;
                        midiIndex--;
                    }
                    else
                    {
                        midiSheet.Cells[midiIndex, 14].Value = "ERROR";
                        return false; //error detected
                    }
                }
                else
                {
                    midiIndex--;
                }
            }
            return true; //No errors found 
        }

        /// <summary>
        /// Convert the entire excel range into an array. Then, scan through the array.
        /// Should an error be detected, you convert 10 items into one big string. You also group 3 notes into another string.
        /// Then, you search for the index of the small group inside the big group. 
        /// </summary>
        /// <param name="midiSheet">The sheet of the sample, representing the playthrough.</param>
        /// <param name="excerptSheet">The excerpt sheet, representing the score.</param>
        /// <returns></returns>
        public bool GroupingDetection(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            List<Node> midiNotes = GetColumnAsList(midiSheet, 7);
            List<Node> excerptNotes = GetColumnAsList(excerptSheet, 2);

            int midiIndex = 0;
            int excerptIndex = 0;

            while(excerptIndex < excerptNotes.Count && midiIndex < midiNotes.Count)
            {
                //Development was not considered for the time scope of the project.
            }
            return false;
        }

        /// <summary>
        /// Gets the specified column as a List.
        /// </summary>
        /// <param name="sheet">The sheet to get the data from.</param>
        /// <param name="col">The column to read.</param>
        /// <returns></returns>
        public List<Node> GetColumnAsList(ExcelWorksheet sheet, int col)
        {
            List<Node> columnList = new List<Node>();

            int lastRow = sheet.Dimension.End.Row;
            int index = 2; //Skip header.

            //Traverse the sheet.
            while(index <= lastRow)
            {
                if(sheet.Cells[index, col].Value != null)   //If the cell is not empty.
                {
                    columnList.Add(new Node(index, sheet.Cells[index, col].Text.Trim().ToUpper()) );  //Add the note to the list.
                }
                index++;
            }
            //Return the list.
            return columnList;
        }

        public bool IsDigitsOnly(string str)
        {
            if(str == null || str.Length == 0)
            {
                return false;
            }
            foreach( char c in str)
            {
                if (c < '0' || c > '9')
                {
                    return false;
                }
            }
            return true;
        }

        public bool CheckNote(string str)
        {
            if(str == null || str.Length < 2 || str.Length > 3)
            {
                return false;
            }
            for(int i = 0; i < str.Length; i++)
            {
                if(i == 0)
                {
                    if((str[i] < 'A' || str[i] > 'G') && (str[i] < 'a' || str[i] > 'g'))
                    {
                        return false;
                    }
                }
                else if(i == 1)
                {
                    if(str[i] < '0' || str[i] > '7') //The notes.json file only goes up to 7
                    {
                        return false;
                    }
                }
                else
                {
                    if(str[i] != '#')
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public class Node
        {
            private int row;        //The row in the sheet where the note was found.
            private string note;    //The note value itself.

            public Node(int row, string note)
            {
                this.row = row;
                this.note = note;
            }

            public int Row
            {
                get { return row; }
                set { row = value; }
            }

            public string Note
            {
                get { return note; }
                set { note = value; }
            }
        }
    }
}
