using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    public interface IIconCatalogDataSource
    {
        string FindManifest();
        string ReadAllText(string path);
        DateTime GetLastWriteTimeUtc(string path);
    }

    public sealed class FileIconCatalogDataSource : IIconCatalogDataSource
    {
        private readonly string _root; private readonly string _relativePath;
        public FileIconCatalogDataSource(string root, string relativePath) { _root=root??throw new ArgumentNullException("root"); _relativePath=relativePath??throw new ArgumentNullException("relativePath"); }
        public string FindManifest() { string runtime=Path.Combine(_root,_relativePath); if(File.Exists(runtime)) return runtime; string dev=Path.GetFullPath(Path.Combine(_root,"..","..",_relativePath)); return File.Exists(dev)?dev:null; }
        public string ReadAllText(string path) { return File.ReadAllText(path); }
        public DateTime GetLastWriteTimeUtc(string path) { return File.GetLastWriteTimeUtc(path); }
    }

    public sealed class IconManifestParser
    {
        public IconCollectionManifest Parse(string json)
        {
            if(string.IsNullOrWhiteSpace(json)) return new IconCollectionManifest();
            using(var stream=new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json)))
                return (IconCollectionManifest)new DataContractJsonSerializer(typeof(IconCollectionManifest)).ReadObject(stream) ?? new IconCollectionManifest();
        }
        public IconCollectionManifest Clone(IconCollectionManifest source)
        {
            if(source==null) return new IconCollectionManifest();
            using(var stream=new MemoryStream()) { var serializer=new DataContractJsonSerializer(typeof(IconCollectionManifest)); serializer.WriteObject(stream,source); stream.Position=0; return (IconCollectionManifest)serializer.ReadObject(stream); }
        }
    }

    public sealed class IconManifestCache
    {
        private readonly object _gate=new object(); private string _path; private DateTime _stamp; private IconCollectionManifest _value;
        public IconCollectionManifest Get(string path, DateTime stamp, Func<string,IconCollectionManifest> loader)
        {
            lock(_gate) { if(_value!=null&&string.Equals(_path,path,StringComparison.OrdinalIgnoreCase)&&_stamp==stamp) return _value; _value=loader(path)??new IconCollectionManifest(); _path=path; _stamp=stamp; return _value; }
        }
        public void Invalidate() { lock(_gate){_value=null;_path=null;_stamp=default(DateTime);} }
    }

    public sealed class IconSearchIndex
    {
        private readonly IReadOnlyList<IconCatalogItem> _icons;
        public IconSearchIndex(IEnumerable<IconCatalogItem> icons) { _icons=(icons??Enumerable.Empty<IconCatalogItem>()).ToList(); }
        public IReadOnlyList<IconCatalogItem> Search(string search=null,string collectionId=null,string theme=null,string style=null,bool favoritesOnly=false)
        {
            IEnumerable<IconCatalogItem> query=_icons;
            if(!string.IsNullOrWhiteSpace(collectionId)) query=query.Where(i=>string.Equals(i.CollectionId,collectionId,StringComparison.OrdinalIgnoreCase));
            if(!string.IsNullOrWhiteSpace(theme)) query=query.Where(i=>string.Equals(i.Theme,theme,StringComparison.OrdinalIgnoreCase));
            if(!string.IsNullOrWhiteSpace(style)) query=query.Where(i=>string.Equals(i.Style,style,StringComparison.OrdinalIgnoreCase));
            if(favoritesOnly) query=query.Where(i=>i.IsFavorite);
            if(!string.IsNullOrWhiteSpace(search)){string needle=search.Trim().ToLowerInvariant();query=query.Where(i=>i.SearchText.Contains(needle));}
            return query.OrderBy(i=>i.Name).ThenBy(i=>i.Style).ToList();
        }
    }
}
