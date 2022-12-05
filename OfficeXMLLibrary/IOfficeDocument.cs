// ConsoleApp6
// OfficeXMLLibrary
// IOfficeDocument.cs
// ---------------------------------------------
// Alex Samarkin
// Alex
// 
// 20:28 04 12 2022

namespace OfficeXMLLibrary
{
    public interface IOfficeDocument
    {
        void Create(string FulllName);
        void Open(string FulllName);
        void Close();
    }
}