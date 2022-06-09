
class Program
{
  static void Main(string[] args)
  {
    do
    {
      Console.Write("\nInput your filepath: ");
      string? filepath = Console.ReadLine(); // "/Users/ryohei/Documents/Blank.docx";
      if (string.IsNullOrEmpty(filepath)) break;
      PracticeLibrary.Class1.OpenWordprocessingDocumentReadonly(filepath);
      Console.WriteLine("\nOne Loop Done!");

    } while (true);
    Console.WriteLine("\nFinished!");

    return;
  }
}
