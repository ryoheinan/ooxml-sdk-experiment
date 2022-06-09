using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace PracticeLibrary;

public class Class1
{
  public static void OpenWordprocessingDocumentReadonly(string filepath)
  {
    // Open a WordprocessingDocument based on a filepath.
    using (WordprocessingDocument wordDocument =
        WordprocessingDocument.Open(filepath, false))
    {
      // Assign a reference to the existing document body.  
      Body? body = wordDocument.MainDocumentPart?.Document.Body;
      if (body == null)
      {
        return;
      }
      foreach (DocumentFormat.OpenXml.OpenXmlElement item in body.Elements())
      {
        Console.WriteLine("\n===⏬ TEXT ⏬===");
        Console.WriteLine(item.InnerText);
        Console.WriteLine("\n===⏬ XML ⏬===");
        Console.WriteLine(item.InnerXml);
      }

      // Console.WriteLine("WHOLE XML\n");
      // Console.WriteLine(body.InnerXml);
      // Console.WriteLine("\n======\n");


      // Console.WriteLine(body.Elements().Count());

      // Attempt to add some text.
      Paragraph para = body.AppendChild(new Paragraph());
      Run run = para.AppendChild(new Run());
      run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

      // Call Save to generate an exception and show that access is read-only.
      // wordDocument.MainDocumentPart.Document.Save();
    }
  }
}
