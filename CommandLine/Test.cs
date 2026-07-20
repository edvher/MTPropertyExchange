/*
* Test class: main testing class
*
* Version:		1.0
*/

using System;
using CommandLine.Utility;

namespace CommandLine{
	/// <summary>
	/// Testing class
	/// </summary>
	class Test{
		/// <summary>
		/// Main loop
		/// </summary>
		[STAThread]
		static void Main(string[] Args){
			// Command line parsing
			Arguments CommandLine=new Arguments(Args);

			// Look for specific arguments values and display them if they exist (return null if they don't)
			if(CommandLine["param1"]!=null) Console.WriteLine("Param1 value: "+CommandLine["param1"]);
			else Console.WriteLine("Param1 not defined !");
			
			if(CommandLine["height"]!=null) Console.WriteLine("Height value: "+CommandLine["height"]);
			else Console.WriteLine("Height not defined !");
			
			if(CommandLine["width"]!=null) Console.WriteLine("Width value: "+CommandLine["width"]);
			else Console.WriteLine("Width not defined !");
			
			if(CommandLine["size"]!=null) Console.WriteLine("Size value: "+CommandLine["size"]);
			else Console.WriteLine("Size not defined !");
			
			if(CommandLine["debug"]!=null) Console.WriteLine("Debug value: "+CommandLine["debug"]);
			else Console.WriteLine("Debug not defined !");

			// Wait for key
			Console.Out.WriteLine("Arguments parsed. Press a key...");
			Console.Read();
			}
		}
	}
