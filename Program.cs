using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Threading;

namespace WordDocumentApp
{
	public class Program
	{
		private static volatile Application app;
		private static volatile Document document;

		private static List<Thread> threadsList;
		private static Semaphore semaphore = new Semaphore(1, 1);
		public static void Main(string[] args)
		{
			threadsList = new List<Thread>()
		{
			new Thread(() =>
			{
				semaphore.WaitOne();

				app = new Application();

				semaphore.Release();
			}),
			new Thread(() =>
			{
				semaphore.WaitOne();

				document = app.Documents.Add();
				var paragraph = document.Paragraphs.Add();
				paragraph.Range.Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.";

				semaphore.Release();
			}),
			new Thread(() =>
			{
				semaphore.WaitOne();

				document.SaveAs2($"{AppDomain.CurrentDomain.BaseDirectory}\\TestFile.doc");
				document.Close();
				app.Quit();

				semaphore.Release();
			})
		};

			threadsList.ForEach(thread =>
			{
				thread.Start();
			});
		}
	}
}
