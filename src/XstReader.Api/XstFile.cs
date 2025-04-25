using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System;
using XstReader.ElementProperties;
using XstReader;

public class XstFile : XstElement, IDisposable
{
    private NDB _Ndb;
    internal new NDB Ndb => _Ndb ?? (_Ndb = new NDB(this));

    private LTP _Ltp;
    internal new LTP Ltp => _Ltp ?? (_Ltp = new LTP(Ndb));

    private string _FileName = null;
    private Stream _InputStream = null;

    /// <summary>
    /// FileName of the .pst or .ost file to read
    /// </summary>
    public string FileName
    {
        get => _FileName;
        set => SetFileName(value);
    }

    private void SetFileName(string fileName)
    {
        _FileName = fileName;
        _InputStream = null; // Clear the stream if a new file is set
        ClearContents();
    }

    private FileStream _ReadStream = null;
    internal FileStream ReadStream
    {
        get
        {
            if (_InputStream != null)
                throw new InvalidOperationException("Cannot access ReadStream when using a custom input stream.");
            return _ReadStream ?? (_ReadStream = new FileStream(FileName, FileMode.Open, FileAccess.Read));
        }
    }

    internal Stream InputStream
    {
        get
        {
            if (_InputStream == null && _ReadStream == null)
                throw new InvalidOperationException("No valid input stream or file stream is available.");
            return _InputStream ?? ReadStream;
        }
    }

    internal object StreamLock { get; } = new object();

    private XstFolder _RootFolder = null;
    /// <summary>
    /// The Root Folder of the XstFile. (Loaded when needed)
    /// </summary>
    public XstFolder RootFolder => _RootFolder ?? (_RootFolder = new XstFolder(this, new NID(EnidSpecial.NID_ROOT_FOLDER)));

    /// <summary>
    /// The Path of this Element
    /// </summary>
    [DisplayName("Path")]
    [Category("General")]
    [Description(@"The Path of this Element")]
    public override string Path => System.IO.Path.GetFileName(this.FileName);

    /// <summary>
    /// The Parents of this Element
    /// </summary>
    [Browsable(false)]
    public override XstElement Parent => null;

    #region Ctor
    /// <summary>
    /// Constructor for file-based input
    /// </summary>
    /// <param name="fileName">The .pst or .ost file to open</param>
    public XstFile(string fileName) : base(XstElementType.File)
    {
        FileName = fileName;
    }

    /// <summary>
    /// Constructor for stream-based input
    /// </summary>
    /// <param name="inputStream">The input stream containing the .pst or .ost data</param>
    public XstFile(Stream inputStream) : base(XstElementType.File)
    {
        _InputStream = inputStream ?? throw new ArgumentNullException(nameof(inputStream));
        _FileName = null; // Clear the file name since we're using a stream
    }
    #endregion Ctor

    private void ClearStream()
    {
        if (_ReadStream != null)
        {
            _ReadStream.Close();
            _ReadStream.Dispose();
            _ReadStream = null;
        }

        if (_InputStream != null)
        {
            _InputStream.Dispose();
            _InputStream = null;
        }
    }

    /// <summary>
    /// Clears information and memory used in RootFolder
    /// </summary>
    private void ClearRootFolder()
    {
        if (_RootFolder != null)
        {
            _RootFolder.ClearContents();
            _RootFolder = null;
        }
    }

    /// <summary>
    /// Clears all information and memory used by the object
    /// </summary>
    public override void ClearContents()
    {
        ClearStream();
        ClearRootFolder();

        _Ndb = null;
        _Ltp = null;
    }

    /// <summary>
    /// Disposes memory used by the object
    /// </summary>
    public void Dispose()
    {
        ClearContents();
    }

    /// <summary>
    /// Gets the String representation of the object
    /// </summary>
    /// <returns></returns>
    public override string ToString()
    {
        return System.IO.Path.GetFileName(FileName ?? "Stream");
    }

    private protected override IEnumerable<XstProperty> LoadProperties()
    {
        return new XstProperty[0];
    }

    private protected override XstProperty LoadProperty(PropertyCanonicalName tag)
    {
        return null;
    }

    private protected override bool CheckProperty(PropertyCanonicalName tag)
    {
        return false;
    }
}
