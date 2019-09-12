﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Midi_Analyzer.Logic
{
    class FileChecker
    {
        /*
         * This class is designed to detect file related errors, mainly in the UI.
         * This can be that a file does not exist, is of invalid type, or the file is open
         * when it should be closed.
         */


        /// <summary>
        /// Checks if a file at a given path exists.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool FileExists(string path)
        {
            return File.Exists(path);
        }

        /// <summary>
        /// Checks if a directory at a given path exists.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool FolderExists(string path)
        {
            return Directory.Exists(path);
        }

        
    }
}
