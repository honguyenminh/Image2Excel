using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Image2Excel
{
    enum ParseResult
    {
        Success,
        Failed,
        Help
    }
    class ArgParser
    {
        public string ImagePath { get; private set; }
        public string OutputPath { get; private set; }

        public ParseResult TryParse(string[] args)
        {
            // Help
            if (args.Contains("-h") || args.Contains("--help"))
            {
                return ParseResult.Help;
            }
            switch (args.Length)
            {
                case 1:
                    // Image path only, auto make output filename
                    ImagePath = args[0];
                    OutputPath = args[0] + ".xlsx";
                    break;
                case 2:
                    // Image and output path given
                    ImagePath = args[0];
                    OutputPath = args[1];
                    break;
                default:
                    return ParseResult.Failed;
            }
            return ParseResult.Success;
        }
    }
}
