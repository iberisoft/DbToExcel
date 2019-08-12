using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DbToExcel
{
    class WorkbookBinding
    {
        public WorkbookBinding(XDocument document)
        {
            Worksheets = document.Root.Elements("Worksheet").Select(element2 => new WorksheetBinding(element2)).ToList();
        }

        public List<WorksheetBinding> Worksheets { get; }
    }

    class WorksheetBinding
    {
        public WorksheetBinding(XElement element)
        {
            Name = (string)element.Attribute(nameof(Name));
            Cells = element.Elements("Cell").Select(element2 => new CellBinding(element2)).ToList();
        }

        public string Name { get; }

        public List<CellBinding> Cells { get; }
    }

    class CellBinding
    {
        public CellBinding(XElement element)
        {
            Name = (string)element.Attribute(nameof(Name));
            Source = (string)element.Attribute(nameof(Source));
            Format = (string)element.Attribute(nameof(Format));
        }

        public string Name { get; }

        public string Source { get; }

        public string Format { get; }
    }
}
