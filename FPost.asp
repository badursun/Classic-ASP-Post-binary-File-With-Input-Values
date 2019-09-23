<%
'####################################
' Custom SQL Query Helper
' Classic ASP & File Post Builder
' 2019  (c) Anthony Burak DURSUN
' 23/09/2019   badursun@gmail.com
' http://www.burakdursun.com
' https://github.com/badursun/Classic-ASP-Post-binary-File-With-Input-Values/
'####################################

Class FilePost
  Dim FilePostArray()
  Dim FilePostErrors()
  Dim FPSize, FPErrSize

  Dim FileTextArray()
  Dim FileTextArraySize

  Private ErrorCount
  Private letPostURL, letFormType
  Private ArrayClearSize, BoundrySize, strDebugBoundry
  Private PostHTTPAnswer, PostHTTPStatus

  '--------------------------------
  ' Class Init
  '--------------------------------
  Private Sub Class_Initialize()
    ArrayClearSize= -1
    BoundrySize   = 20

    FPSize              = -1
    FPErrSize           = -1
    FileTextArraySize   = -1
    ErrorCount          = 0
    letPostURL          = Null
    letFormType         = Null
    PostHTTPAnswer      = ""
    PostHTTPStatus      = ""
    strDebugBoundry     = ""
  End Sub

  '--------------------------------
  ' Class Destroy
  '--------------------------------
  Private Sub Class_Terminate()
    ReDim PRESERVE FilePostArray( ArrayClearSize )
    ReDim PRESERVE FilePostErrors( ArrayClearSize )
  End Sub

  '--------------------------------
  ' Add File
  '--------------------------------
  Public Sub AddFile(vFileName)
    If FileExist(vFileName) = True Then
      ReDim PRESERVE FilePostArray( FPSize )
      FPSize=FPSize+1

      ReDim PRESERVE FilePostArray( FPSize )
      FilePostArray(FPSize) = vFileName
    End IF
  End Sub

  '--------------------------------
  ' Add Text
  '--------------------------------
  Public Sub AddText(vInputName, vFileName)
      ReDim PRESERVE FileTextArray( FileTextArraySize )
      FileTextArraySize=FileTextArraySize+1

      ReDim PRESERVE FileTextArray( FileTextArraySize )
      FileTextArray(FileTextArraySize) = ""&vInputName&"=|*|="&vFileName&""
  End Sub

  '--------------------------------
  ' File Exits
  '--------------------------------
  Private Function FileExist(vFileName)
    If Len(vFileName) < 4 Then 
        Call ErrorRaise(1, vFileName)
        FileExist = False
        Exit Function
    End If

    Dim FileExist_FSO
    Set FileExist_FSO = Server.CreateObject("Scripting.FileSystemObject")
      If FileExist_FSO.FileExists( Server.MapPath(vFileName) ) Then
        FileExist = True
      Else
        Call ErrorRaise(2, vFileName)
        FileExist = False
      End if
    Set FileExist_FSO = Nothing
  End Function

  '--------------------------------
  ' Error Management
  '--------------------------------
  Private Sub ErrorRaise(vErrCode, vDescription)
    ErrorCount=ErrorCount+1

    Select Case vErrCode
      Case 1 
        strErrorRaise = "Enter a valid file name for "&(FPSize+2)&".AddFile ("""&vDescription&""")"
      Case 2 
        strErrorRaise =  "File not found in specified path for "&(FPSize+2)&".AddFile ("""&vDescription&""")"
      Case 3
        strErrorRaise = ""
      Case 4
        strErrorRaise = "Please LET Post URL Address."
      Case Else 
        strErrorRaise = "An undefined error occurred for "&(FPSize+2)&".AddFile ("""&vDescription&""")"
    End Select

    ReDim PRESERVE FilePostErrors( FPErrSize )
    FPErrSize=FPErrSize+1

    ReDim PRESERVE FilePostErrors( FPErrSize )
    FilePostErrors(FPErrSize) = strErrorRaise

    Response.Write strErrorRaise & "<br />"
  End Sub

  '--------------------------------
  ' Debug
  '--------------------------------
  Public Sub Debug()
    'Debug Files
    Response.Write "<h5>Added Files</h5>"
    For i=0 To UBound(FilePostArray)
      Response.Write  "<strong>ERROR: </strong> "&i&" - " & FilePostArray(i) & "<br />"
    Next
    'Error Debug
    Response.Write "<h5>Raised Errors</h5>"
    For i=0 To UBound(FilePostErrors)
      Response.Write  "<strong>ERROR: </strong> "&i&" - " & FilePostErrors(i) & "<br />"
    Next

  End Sub

  '--------------------------------
  ' Binary To String Converter
  '--------------------------------
  Public Function BinaryToString(Binary)
    Dim I, S
    For I = 1 To LenB(Binary)
      S = S & Chr(AscB(MidB(Binary, I, 1)))
    Next
    BinaryToString = S
  End Function

  '--------------------------------
  ' Read Binary File
  '--------------------------------
  Private Function ReadBinaryFile(FileName)
    Dim BinaryStream
    Set BinaryStream = CreateObject("ADODB.Stream")
    'Specify stream type - we want To get binary data.
    BinaryStream.Type = 1
    'Open the stream
    BinaryStream.Open
    'Load the file data from disk To stream object
    BinaryStream.LoadFromFile FileName
    'Open the stream And get binary data from the object
    ReadBinaryFile = BinaryStream.Read
  End Function

  '--------------------------------
  ' Stream String To Binary
  '--------------------------------
  Function Stream_StringToBinary(Text, CharSet)
    Dim BinaryStream 'As New Stream
    Set BinaryStream = CreateObject("ADODB.Stream")  
      BinaryStream.Type = 2
      If Len(CharSet) > 0 Then
        BinaryStream.CharSet  = CharSet
      Else
        BinaryStream.CharSet  = "us-ascii"
      End If
      BinaryStream.Open
      BinaryStream.WriteText Text
      BinaryStream.Position   = 0
      BinaryStream.Type       = 1
      BinaryStream.Position   = 0
      Stream_StringToBinary = BinaryStream.Read
    Set BinaryStream = Nothing
  End Function

  '--------------------------------
  ' Random Boundry Generator
  '--------------------------------
  Private Function RandomBoundryGenerator()
    Const szDefault = "abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
    Dim nCount, sRet, nNumber, nLength
    Randomize
    sValidChars = szDefault   
    nLength = Len( sValidChars )
    For nCount = 1 To BoundrySize
      nNumber = Int((nLength * Rnd) + 1)
      sRet = sRet & Mid( sValidChars, nNumber, 1 )
    Next
    RandomBoundryGenerator = sRet  
  End Function

  '--------------------------------
  ' MimeType & Extentions
  '--------------------------------
  Private Function MimeTypes(vFileExtention)
    '----------------------------------------------------
    ' https://www.freeformatter.com/mime-types-list.html
    '----------------------------------------------------  
    Select Case GetFileExtention(vFileExtention)
      Case "x3d" : MimeTypes = "application/vnd.hzn-3d-crossword" 
      Case "3gp" : MimeTypes = "video/3gpp" 
      Case "3g2" : MimeTypes = "video/3gpp2" 
      Case "mseq" : MimeTypes = "application/vnd.mseq" 
      Case "pwn" : MimeTypes = "application/vnd.3m.post-it-notes" 
      Case "plb" : MimeTypes = "application/vnd.3gpp.pic-bw-large" 
      Case "psb" : MimeTypes = "application/vnd.3gpp.pic-bw-small" 
      Case "pvb" : MimeTypes = "application/vnd.3gpp.pic-bw-var" 
      Case "tcap" : MimeTypes = "application/vnd.3gpp2.tcap" 
      Case "7z" : MimeTypes = "application/x-7z-compressed" 
      Case "abw" : MimeTypes = "application/x-abiword" 
      Case "ace" : MimeTypes = "application/x-ace-compressed" 
      Case "acc" : MimeTypes = "application/vnd.americandynamics.acc" 
      Case "acu" : MimeTypes = "application/vnd.acucobol" 
      Case "atc" : MimeTypes = "application/vnd.acucorp" 
      Case "adp" : MimeTypes = "audio/adpcm" 
      Case "aab" : MimeTypes = "application/x-authorware-bin" 
      Case "aam" : MimeTypes = "application/x-authorware-map" 
      Case "aas" : MimeTypes = "application/x-authorware-seg" 
      Case "air" : MimeTypes = "application/vnd.adobe.air-application-installer-package+zip" 
      Case "swf" : MimeTypes = "application/x-shockwave-flash" 
      Case "fxp" : MimeTypes = "application/vnd.adobe.fxp" 
      Case "pdf" : MimeTypes = "application/pdf" 
      Case "ppd" : MimeTypes = "application/vnd.cups-ppd" 
      Case "dir" : MimeTypes = "application/x-director" 
      Case "xdp" : MimeTypes = "application/vnd.adobe.xdp+xml" 
      Case "xfdf" : MimeTypes = "application/vnd.adobe.xfdf" 
      Case "aac" : MimeTypes = "audio/x-aac" 
      Case "ahead" : MimeTypes = "application/vnd.ahead.space" 
      Case "azf" : MimeTypes = "application/vnd.airzip.filesecure.azf" 
      Case "azs" : MimeTypes = "application/vnd.airzip.filesecure.azs" 
      Case "azw" : MimeTypes = "application/vnd.amazon.ebook" 
      Case "ami" : MimeTypes = "application/vnd.amiga.ami" 
      Case "apk" : MimeTypes = "application/vnd.android.package-archive" 
      Case "cii" : MimeTypes = "application/vnd.anser-web-certificate-issue-initiation" 
      Case "fti" : MimeTypes = "application/vnd.anser-web-funds-transfer-initiation" 
      Case "atx" : MimeTypes = "application/vnd.antix.game-component" 
      Case "dmg" : MimeTypes = "application/x-apple-diskimage" 
      Case "mpkg" : MimeTypes = "application/vnd.apple.installer+xml" 
      Case "aw" : MimeTypes = "application/applixware" 
      Case "les" : MimeTypes = "application/vnd.hhe.lesson-player" 
      Case "swi" : MimeTypes = "application/vnd.aristanetworks.swi" 
      Case "s" : MimeTypes = "text/x-asm" 
      Case "atomcat" : MimeTypes = "application/atomcat+xml" 
      Case "atomsvc" : MimeTypes = "application/atomsvc+xml" 
      Case "atom, .xml" : MimeTypes = "application/atom+xml" 
      Case "ac" : MimeTypes = "application/pkix-attr-cert" 
      Case "aif" : MimeTypes = "audio/x-aiff" 
      Case "avi" : MimeTypes = "video/x-msvideo" 
      Case "aep" : MimeTypes = "application/vnd.audiograph" 
      Case "dxf" : MimeTypes = "image/vnd.dxf" 
      Case "dwf" : MimeTypes = "model/vnd.dwf" 
      Case "par" : MimeTypes = "text/plain-bas" 
      Case "bcpio" : MimeTypes = "application/x-bcpio" 
      Case "bin" : MimeTypes = "application/octet-stream" 
      Case "bmp" : MimeTypes = "image/bmp" 
      Case "torrent" : MimeTypes = "application/x-bittorrent" 
      Case "cod" : MimeTypes = "application/vnd.rim.cod" 
      Case "mpm" : MimeTypes = "application/vnd.blueice.multipass" 
      Case "bmi" : MimeTypes = "application/vnd.bmi" 
      Case "sh" : MimeTypes = "application/x-sh" 
      Case "btif" : MimeTypes = "image/prs.btif" 
      Case "rep" : MimeTypes = "application/vnd.businessobjects" 
      Case "bz" : MimeTypes = "application/x-bzip" 
      Case "bz2" : MimeTypes = "application/x-bzip2" 
      Case "csh" : MimeTypes = "application/x-csh" 
      Case "c" : MimeTypes = "text/x-c" 
      Case "cdxml" : MimeTypes = "application/vnd.chemdraw+xml" 
      Case "css" : MimeTypes = "text/css" 
      Case "cdx" : MimeTypes = "chemical/x-cdx" 
      Case "cml" : MimeTypes = "chemical/x-cml" 
      Case "csml" : MimeTypes = "chemical/x-csml" 
      Case "cdbcmsg" : MimeTypes = "application/vnd.contact.cmsg" 
      Case "cla" : MimeTypes = "application/vnd.claymore" 
      Case "c4g" : MimeTypes = "application/vnd.clonk.c4group" 
      Case "sub" : MimeTypes = "image/vnd.dvb.subtitle" 
      Case "cdmia" : MimeTypes = "application/cdmi-capability" 
      Case "cdmic" : MimeTypes = "application/cdmi-container" 
      Case "cdmid" : MimeTypes = "application/cdmi-domain" 
      Case "cdmio" : MimeTypes = "application/cdmi-object" 
      Case "cdmiq" : MimeTypes = "application/cdmi-queue" 
      Case "c11amc" : MimeTypes = "application/vnd.cluetrust.cartomobile-config" 
      Case "c11amz" : MimeTypes = "application/vnd.cluetrust.cartomobile-config-pkg" 
      Case "ras" : MimeTypes = "image/x-cmu-raster" 
      Case "dae" : MimeTypes = "model/vnd.collada+xml" 
      Case "csv" : MimeTypes = "text/csv" 
      Case "cpt" : MimeTypes = "application/mac-compactpro" 
      Case "wmlc" : MimeTypes = "application/vnd.wap.wmlc" 
      Case "cgm" : MimeTypes = "image/cgm" 
      Case "ice" : MimeTypes = "x-conference/x-cooltalk" 
      Case "cmx" : MimeTypes = "image/x-cmx" 
      Case "xar" : MimeTypes = "application/vnd.xara" 
      Case "cmc" : MimeTypes = "application/vnd.cosmocaller" 
      Case "cpio" : MimeTypes = "application/x-cpio" 
      Case "clkx" : MimeTypes = "application/vnd.crick.clicker" 
      Case "clkk" : MimeTypes = "application/vnd.crick.clicker.keyboard" 
      Case "clkp" : MimeTypes = "application/vnd.crick.clicker.palette" 
      Case "clkt" : MimeTypes = "application/vnd.crick.clicker.template" 
      Case "clkw" : MimeTypes = "application/vnd.crick.clicker.wordbank" 
      Case "wbs" : MimeTypes = "application/vnd.criticaltools.wbs+xml" 
      Case "cryptonote" : MimeTypes = "application/vnd.rig.cryptonote" 
      Case "cif" : MimeTypes = "chemical/x-cif" 
      Case "cmdf" : MimeTypes = "chemical/x-cmdf" 
      Case "cu" : MimeTypes = "application/cu-seeme" 
      Case "cww" : MimeTypes = "application/prs.cww" 
      Case "curl" : MimeTypes = "text/vnd.curl" 
      Case "dcurl" : MimeTypes = "text/vnd.curl.dcurl" 
      Case "mcurl" : MimeTypes = "text/vnd.curl.mcurl" 
      Case "scurl" : MimeTypes = "text/vnd.curl.scurl" 
      Case "car" : MimeTypes = "application/vnd.curl.car" 
      Case "pcurl" : MimeTypes = "application/vnd.curl.pcurl" 
      Case "cmp" : MimeTypes = "application/vnd.yellowriver-custom-menu" 
      Case "dssc" : MimeTypes = "application/dssc+der" 
      Case "xdssc" : MimeTypes = "application/dssc+xml" 
      Case "deb" : MimeTypes = "application/x-debian-package" 
      Case "uva" : MimeTypes = "audio/vnd.dece.audio" 
      Case "uvi" : MimeTypes = "image/vnd.dece.graphic" 
      Case "uvh" : MimeTypes = "video/vnd.dece.hd" 
      Case "uvm" : MimeTypes = "video/vnd.dece.mobile" 
      Case "uvu" : MimeTypes = "video/vnd.uvvu.mp4" 
      Case "uvp" : MimeTypes = "video/vnd.dece.pd" 
      Case "uvs" : MimeTypes = "video/vnd.dece.sd" 
      Case "uvv" : MimeTypes = "video/vnd.dece.video" 
      Case "dvi" : MimeTypes = "application/x-dvi" 
      Case "seed" : MimeTypes = "application/vnd.fdsn.seed" 
      Case "dtb" : MimeTypes = "application/x-dtbook+xml" 
      Case "res" : MimeTypes = "application/x-dtbresource+xml" 
      Case "ait" : MimeTypes = "application/vnd.dvb.ait" 
      Case "svc" : MimeTypes = "application/vnd.dvb.service" 
      Case "eol" : MimeTypes = "audio/vnd.digital-winds" 
      Case "djvu" : MimeTypes = "image/vnd.djvu" 
      Case "dtd" : MimeTypes = "application/xml-dtd" 
      Case "mlp" : MimeTypes = "application/vnd.dolby.mlp" 
      Case "wad" : MimeTypes = "application/x-doom" 
      Case "dpg" : MimeTypes = "application/vnd.dpgraph" 
      Case "dra" : MimeTypes = "audio/vnd.dra" 
      Case "dfac" : MimeTypes = "application/vnd.dreamfactory" 
      Case "dts" : MimeTypes = "audio/vnd.dts" 
      Case "dtshd" : MimeTypes = "audio/vnd.dts.hd" 
      Case "dwg" : MimeTypes = "image/vnd.dwg" 
      Case "geo" : MimeTypes = "application/vnd.dynageo" 
      Case "es" : MimeTypes = "application/ecmascript" 
      Case "mag" : MimeTypes = "application/vnd.ecowin.chart" 
      Case "mmr" : MimeTypes = "image/vnd.fujixerox.edmics-mmr" 
      Case "rlc" : MimeTypes = "image/vnd.fujixerox.edmics-rlc" 
      Case "exi" : MimeTypes = "application/exi" 
      Case "mgz" : MimeTypes = "application/vnd.proteus.magazine" 
      Case "epub" : MimeTypes = "application/epub+zip" 
      Case "eml" : MimeTypes = "message/rfc822" 
      Case "nml" : MimeTypes = "application/vnd.enliven" 
      Case "xpr" : MimeTypes = "application/vnd.is-xpr" 
      Case "xif" : MimeTypes = "image/vnd.xiff" 
      Case "xfdl" : MimeTypes = "application/vnd.xfdl" 
      Case "emma" : MimeTypes = "application/emma+xml" 
      Case "ez2" : MimeTypes = "application/vnd.ezpix-album" 
      Case "ez3" : MimeTypes = "application/vnd.ezpix-package" 
      Case "fst" : MimeTypes = "image/vnd.fst" 
      Case "fvt" : MimeTypes = "video/vnd.fvt" 
      Case "fbs" : MimeTypes = "image/vnd.fastbidsheet" 
      Case "fe_launch" : MimeTypes = "application/vnd.denovo.fcselayout-link" 
      Case "f4v" : MimeTypes = "video/x-f4v" 
      Case "flv" : MimeTypes = "video/x-flv" 
      Case "fpx" : MimeTypes = "image/vnd.fpx" 
      Case "npx" : MimeTypes = "image/vnd.net-fpx" 
      Case "flx" : MimeTypes = "text/vnd.fmi.flexstor" 
      Case "fli" : MimeTypes = "video/x-fli" 
      Case "ftc" : MimeTypes = "application/vnd.fluxtime.clip" 
      Case "fdf" : MimeTypes = "application/vnd.fdf" 
      Case "f" : MimeTypes = "text/x-fortran" 
      Case "mif" : MimeTypes = "application/vnd.mif" 
      Case "fm" : MimeTypes = "application/vnd.framemaker" 
      Case "fh" : MimeTypes = "image/x-freehand" 
      Case "fsc" : MimeTypes = "application/vnd.fsc.weblaunch" 
      Case "fnc" : MimeTypes = "application/vnd.frogans.fnc" 
      Case "ltf" : MimeTypes = "application/vnd.frogans.ltf" 
      Case "ddd" : MimeTypes = "application/vnd.fujixerox.ddd" 
      Case "xdw" : MimeTypes = "application/vnd.fujixerox.docuworks" 
      Case "xbd" : MimeTypes = "application/vnd.fujixerox.docuworks.binder" 
      Case "oas" : MimeTypes = "application/vnd.fujitsu.oasys" 
      Case "oa2" : MimeTypes = "application/vnd.fujitsu.oasys2" 
      Case "oa3" : MimeTypes = "application/vnd.fujitsu.oasys3" 
      Case "fg5" : MimeTypes = "application/vnd.fujitsu.oasysgp" 
      Case "bh2" : MimeTypes = "application/vnd.fujitsu.oasysprs" 
      Case "spl" : MimeTypes = "application/x-futuresplash" 
      Case "fzs" : MimeTypes = "application/vnd.fuzzysheet" 
      Case "g3" : MimeTypes = "image/g3fax" 
      Case "gmx" : MimeTypes = "application/vnd.gmx" 
      Case "gtw" : MimeTypes = "model/vnd.gtw" 
      Case "txd" : MimeTypes = "application/vnd.genomatix.tuxedo" 
      Case "ggb" : MimeTypes = "application/vnd.geogebra.file" 
      Case "ggt" : MimeTypes = "application/vnd.geogebra.tool" 
      Case "gdl" : MimeTypes = "model/vnd.gdl" 
      Case "gex" : MimeTypes = "application/vnd.geometry-explorer" 
      Case "gxt" : MimeTypes = "application/vnd.geonext" 
      Case "g2w" : MimeTypes = "application/vnd.geoplan" 
      Case "g3w" : MimeTypes = "application/vnd.geospace" 
      Case "gsf" : MimeTypes = "application/x-font-ghostscript" 
      Case "bdf" : MimeTypes = "application/x-font-bdf" 
      Case "gtar" : MimeTypes = "application/x-gtar" 
      Case "texinfo" : MimeTypes = "application/x-texinfo" 
      Case "gnumeric" : MimeTypes = "application/x-gnumeric" 
      Case "kml" : MimeTypes = "application/vnd.google-earth.kml+xml" 
      Case "kmz" : MimeTypes = "application/vnd.google-earth.kmz" 
      Case "gqf" : MimeTypes = "application/vnd.grafeq" 
      Case "gif" : MimeTypes = "image/gif" 
      Case "gv" : MimeTypes = "text/vnd.graphviz" 
      Case "gac" : MimeTypes = "application/vnd.groove-account" 
      Case "ghf" : MimeTypes = "application/vnd.groove-help" 
      Case "gim" : MimeTypes = "application/vnd.groove-identity-message" 
      Case "grv" : MimeTypes = "application/vnd.groove-injector" 
      Case "gtm" : MimeTypes = "application/vnd.groove-tool-message" 
      Case "tpl" : MimeTypes = "application/vnd.groove-tool-template" 
      Case "vcg" : MimeTypes = "application/vnd.groove-vcard" 
      Case "h261" : MimeTypes = "video/h261" 
      Case "h263" : MimeTypes = "video/h263" 
      Case "h264" : MimeTypes = "video/h264" 
      Case "hpid" : MimeTypes = "application/vnd.hp-hpid" 
      Case "hps" : MimeTypes = "application/vnd.hp-hps" 
      Case "hdf" : MimeTypes = "application/x-hdf" 
      Case "rip" : MimeTypes = "audio/vnd.rip" 
      Case "hbci" : MimeTypes = "application/vnd.hbci" 
      Case "jlt" : MimeTypes = "application/vnd.hp-jlyt" 
      Case "pcl" : MimeTypes = "application/vnd.hp-pcl" 
      Case "hpgl" : MimeTypes = "application/vnd.hp-hpgl" 
      Case "hvs" : MimeTypes = "application/vnd.yamaha.hv-script" 
      Case "hvd" : MimeTypes = "application/vnd.yamaha.hv-dic" 
      Case "hvp" : MimeTypes = "application/vnd.yamaha.hv-voice" 
      Case "sfd-hdstx" : MimeTypes = "application/vnd.hydrostatix.sof-data" 
      Case "stk" : MimeTypes = "application/hyperstudio" 
      Case "hal" : MimeTypes = "application/vnd.hal+xml" 
      Case "html" : MimeTypes = "text/html" 
      Case "irm" : MimeTypes = "application/vnd.ibm.rights-management" 
      Case "sc" : MimeTypes = "application/vnd.ibm.secure-container" 
      Case "ics" : MimeTypes = "text/calendar" 
      Case "icc" : MimeTypes = "application/vnd.iccprofile" 
      Case "ico" : MimeTypes = "image/x-icon" 
      Case "igl" : MimeTypes = "application/vnd.igloader" 
      Case "ief" : MimeTypes = "image/ief" 
      Case "ivp" : MimeTypes = "application/vnd.immervision-ivp" 
      Case "ivu" : MimeTypes = "application/vnd.immervision-ivu" 
      Case "rif" : MimeTypes = "application/reginfo+xml" 
      Case "3dml" : MimeTypes = "text/vnd.in3d.3dml" 
      Case "spot" : MimeTypes = "text/vnd.in3d.spot" 
      Case "igs" : MimeTypes = "model/iges" 
      Case "i2g" : MimeTypes = "application/vnd.intergeo" 
      Case "cdy" : MimeTypes = "application/vnd.cinderella" 
      Case "xpw" : MimeTypes = "application/vnd.intercon.formnet" 
      Case "fcs" : MimeTypes = "application/vnd.isac.fcs" 
      Case "ipfix" : MimeTypes = "application/ipfix" 
      Case "cer" : MimeTypes = "application/pkix-cert" 
      Case "pki" : MimeTypes = "application/pkixcmp" 
      Case "crl" : MimeTypes = "application/pkix-crl" 
      Case "pkipath" : MimeTypes = "application/pkix-pkipath" 
      Case "igm" : MimeTypes = "application/vnd.insors.igm" 
      Case "rcprofile" : MimeTypes = "application/vnd.ipunplugged.rcprofile" 
      Case "irp" : MimeTypes = "application/vnd.irepository.package+xml" 
      Case "jad" : MimeTypes = "text/vnd.sun.j2me.app-descriptor" 
      Case "jar" : MimeTypes = "application/java-archive" 
      Case "class" : MimeTypes = "application/java-vm" 
      Case "jnlp" : MimeTypes = "application/x-java-jnlp-file" 
      Case "ser" : MimeTypes = "application/java-serialized-object" 
      Case "java" : MimeTypes = "text/x-java-source,java" 
      Case "js" : MimeTypes = "application/javascript" 
      Case "json" : MimeTypes = "application/json" 
      Case "joda" : MimeTypes = "application/vnd.joost.joda-archive" 
      Case "jpm" : MimeTypes = "video/jpm" 
      Case "jpeg" : MimeTypes = "image/jpeg" 
      Case "jpg"        : MimeTypes = "image/jpeg" 
      Case "pjpeg" : MimeTypes = "image/pjpeg" 
      Case "jpgv" : MimeTypes = "video/jpeg" 
      Case "ktz" : MimeTypes = "application/vnd.kahootz" 
      Case "mmd" : MimeTypes = "application/vnd.chipnuts.karaoke-mmd" 
      Case "karbon" : MimeTypes = "application/vnd.kde.karbon" 
      Case "chrt" : MimeTypes = "application/vnd.kde.kchart" 
      Case "kfo" : MimeTypes = "application/vnd.kde.kformula" 
      Case "flw" : MimeTypes = "application/vnd.kde.kivio" 
      Case "kon" : MimeTypes = "application/vnd.kde.kontour" 
      Case "kpr" : MimeTypes = "application/vnd.kde.kpresenter" 
      Case "ksp" : MimeTypes = "application/vnd.kde.kspread" 
      Case "kwd" : MimeTypes = "application/vnd.kde.kword" 
      Case "htke" : MimeTypes = "application/vnd.kenameaapp" 
      Case "kia" : MimeTypes = "application/vnd.kidspiration" 
      Case "kne" : MimeTypes = "application/vnd.kinar" 
      Case "sse" : MimeTypes = "application/vnd.kodak-descriptor" 
      Case "lasxml" : MimeTypes = "application/vnd.las.las+xml" 
      Case "latex" : MimeTypes = "application/x-latex" 
      Case "lbd" : MimeTypes = "application/vnd.llamagraphics.life-balance.desktop" 
      Case "lbe" : MimeTypes = "application/vnd.llamagraphics.life-balance.exchange+xml" 
      Case "jam" : MimeTypes = "application/vnd.jam" 
      Case "apr" : MimeTypes = "application/vnd.lotus-approach" 
      Case "pre" : MimeTypes = "application/vnd.lotus-freelance" 
      Case "nsf" : MimeTypes = "application/vnd.lotus-notes" 
      Case "org" : MimeTypes = "application/vnd.lotus-organizer" 
      Case "scm" : MimeTypes = "application/vnd.lotus-screencam" 
      Case "lwp" : MimeTypes = "application/vnd.lotus-wordpro" 
      Case "lvp" : MimeTypes = "audio/vnd.lucent.voice" 
      Case "m3u" : MimeTypes = "audio/x-mpegurl" 
      Case "m4v" : MimeTypes = "video/x-m4v" 
      Case "hqx" : MimeTypes = "application/mac-binhex40" 
      Case "portpkg" : MimeTypes = "application/vnd.macports.portpkg" 
      Case "mgp" : MimeTypes = "application/vnd.osgeo.mapguide.package" 
      Case "mrc" : MimeTypes = "application/marc" 
      Case "mrcx" : MimeTypes = "application/marcxml+xml" 
      Case "mxf" : MimeTypes = "application/mxf" 
      Case "nbp" : MimeTypes = "application/vnd.wolfram.player" 
      Case "ma" : MimeTypes = "application/mathematica" 
      Case "mathml" : MimeTypes = "application/mathml+xml" 
      Case "mbox" : MimeTypes = "application/mbox" 
      Case "mc1" : MimeTypes = "application/vnd.medcalcdata" 
      Case "mscml" : MimeTypes = "application/mediaservercontrol+xml" 
      Case "cdkey" : MimeTypes = "application/vnd.mediastation.cdkey" 
      Case "mwf" : MimeTypes = "application/vnd.mfer" 
      Case "mfm" : MimeTypes = "application/vnd.mfmp" 
      Case "msh" : MimeTypes = "model/mesh" 
      Case "mads" : MimeTypes = "application/mads+xml" 
      Case "mets" : MimeTypes = "application/mets+xml" 
      Case "mods" : MimeTypes = "application/mods+xml" 
      Case "meta4" : MimeTypes = "application/metalink4+xml" 
      Case "mcd" : MimeTypes = "application/vnd.mcd" 
      Case "flo" : MimeTypes = "application/vnd.micrografx.flo" 
      Case "igx" : MimeTypes = "application/vnd.micrografx.igx" 
      Case "es3" : MimeTypes = "application/vnd.eszigno3+xml" 
      Case "mdb" : MimeTypes = "application/x-msaccess" 
      Case "asf" : MimeTypes = "video/x-ms-asf" 
      Case "exe" : MimeTypes = "application/x-msdownload" 
      Case "cil" : MimeTypes = "application/vnd.ms-artgalry" 
      Case "cab" : MimeTypes = "application/vnd.ms-cab-compressed" 
      Case "ims" : MimeTypes = "application/vnd.ms-ims" 
      Case "application" : MimeTypes = "application/x-ms-application" 
      Case "clp" : MimeTypes = "application/x-msclip" 
      Case "mdi" : MimeTypes = "image/vnd.ms-modi" 
      Case "eot" : MimeTypes = "application/vnd.ms-fontobject" 
      Case "xls" : MimeTypes = "application/vnd.ms-excel" 
      Case "xlam" : MimeTypes = "application/vnd.ms-excel.addin.macroenabled.12" 
      Case "xlsb" : MimeTypes = "application/vnd.ms-excel.sheet.binary.macroenabled.12" 
      Case "xltm" : MimeTypes = "application/vnd.ms-excel.template.macroenabled.12" 
      Case "xlsm" : MimeTypes = "application/vnd.ms-excel.sheet.macroenabled.12" 
      Case "chm" : MimeTypes = "application/vnd.ms-htmlhelp" 
      Case "crd" : MimeTypes = "application/x-mscardfile" 
      Case "lrm" : MimeTypes = "application/vnd.ms-lrm" 
      Case "mvb" : MimeTypes = "application/x-msmediaview" 
      Case "mny" : MimeTypes = "application/x-msmoney" 
      Case "pptx" : MimeTypes = "application/vnd.openxmlformats-officedocument.presentationml.presentation" 
      Case "sldx" : MimeTypes = "application/vnd.openxmlformats-officedocument.presentationml.slide" 
      Case "ppsx" : MimeTypes = "application/vnd.openxmlformats-officedocument.presentationml.slideshow" 
      Case "potx" : MimeTypes = "application/vnd.openxmlformats-officedocument.presentationml.template" 
      Case "xlsx" : MimeTypes = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
      Case "xltx" : MimeTypes = "application/vnd.openxmlformats-officedocument.spreadsheetml.template" 
      Case "docx" : MimeTypes = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" 
      Case "dotx" : MimeTypes = "application/vnd.openxmlformats-officedocument.wordprocessingml.template" 
      Case "obd" : MimeTypes = "application/x-msbinder" 
      Case "thmx" : MimeTypes = "application/vnd.ms-officetheme" 
      Case "onetoc" : MimeTypes = "application/onenote" 
      Case "pya" : MimeTypes = "audio/vnd.ms-playready.media.pya" 
      Case "pyv" : MimeTypes = "video/vnd.ms-playready.media.pyv" 
      Case "ppt" : MimeTypes = "application/vnd.ms-powerpoint" 
      Case "ppam" : MimeTypes = "application/vnd.ms-powerpoint.addin.macroenabled.12" 
      Case "sldm" : MimeTypes = "application/vnd.ms-powerpoint.slide.macroenabled.12" 
      Case "pptm" : MimeTypes = "application/vnd.ms-powerpoint.presentation.macroenabled.12" 
      Case "ppsm" : MimeTypes = "application/vnd.ms-powerpoint.slideshow.macroenabled.12" 
      Case "potm" : MimeTypes = "application/vnd.ms-powerpoint.template.macroenabled.12" 
      Case "mpp" : MimeTypes = "application/vnd.ms-project" 
      Case "pub" : MimeTypes = "application/x-mspublisher" 
      Case "scd" : MimeTypes = "application/x-msschedule" 
      Case "xap" : MimeTypes = "application/x-silverlight-app" 
      Case "stl" : MimeTypes = "application/vnd.ms-pki.stl" 
      Case "cat" : MimeTypes = "application/vnd.ms-pki.seccat" 
      Case "vsd" : MimeTypes = "application/vnd.visio" 
      Case "vsdx" : MimeTypes = "application/vnd.visio2013" 
      Case "wm" : MimeTypes = "video/x-ms-wm" 
      Case "wma" : MimeTypes = "audio/x-ms-wma" 
      Case "wax" : MimeTypes = "audio/x-ms-wax" 
      Case "wmx" : MimeTypes = "video/x-ms-wmx" 
      Case "wmd" : MimeTypes = "application/x-ms-wmd" 
      Case "wpl" : MimeTypes = "application/vnd.ms-wpl" 
      Case "wmz" : MimeTypes = "application/x-ms-wmz" 
      Case "wmv" : MimeTypes = "video/x-ms-wmv" 
      Case "wvx" : MimeTypes = "video/x-ms-wvx" 
      Case "wmf" : MimeTypes = "application/x-msmetafile" 
      Case "trm" : MimeTypes = "application/x-msterminal" 
      Case "doc" : MimeTypes = "application/msword" 
      Case "docm" : MimeTypes = "application/vnd.ms-word.document.macroenabled.12" 
      Case "dotm" : MimeTypes = "application/vnd.ms-word.template.macroenabled.12" 
      Case "wri" : MimeTypes = "application/x-mswrite" 
      Case "wps" : MimeTypes = "application/vnd.ms-works" 
      Case "xbap" : MimeTypes = "application/x-ms-xbap" 
      Case "xps" : MimeTypes = "application/vnd.ms-xpsdocument" 
      Case "mid" : MimeTypes = "audio/midi" 
      Case "mpy" : MimeTypes = "application/vnd.ibm.minipay" 
      Case "afp" : MimeTypes = "application/vnd.ibm.modcap" 
      Case "rms" : MimeTypes = "application/vnd.jcp.javame.midlet-rms" 
      Case "tmo" : MimeTypes = "application/vnd.tmobile-livetv" 
      Case "prc" : MimeTypes = "application/x-mobipocket-ebook" 
      Case "mbk" : MimeTypes = "application/vnd.mobius.mbk" 
      Case "dis" : MimeTypes = "application/vnd.mobius.dis" 
      Case "plc" : MimeTypes = "application/vnd.mobius.plc" 
      Case "mqy" : MimeTypes = "application/vnd.mobius.mqy" 
      Case "msl" : MimeTypes = "application/vnd.mobius.msl" 
      Case "txf" : MimeTypes = "application/vnd.mobius.txf" 
      Case "daf" : MimeTypes = "application/vnd.mobius.daf" 
      Case "fly" : MimeTypes = "text/vnd.fly" 
      Case "mpc" : MimeTypes = "application/vnd.mophun.certificate" 
      Case "mpn" : MimeTypes = "application/vnd.mophun.application" 
      Case "mj2" : MimeTypes = "video/mj2" 
      Case "mpga" : MimeTypes = "audio/mpeg" 
      Case "mxu" : MimeTypes = "video/vnd.mpegurl" 
      Case "mpeg" : MimeTypes = "video/mpeg" 
      Case "m21" : MimeTypes = "application/mp21" 
      Case "mp4a" : MimeTypes = "audio/mp4" 
      Case "mp4" : MimeTypes = "video/mp4" 
      Case "mp4" : MimeTypes = "application/mp4" 
      Case "m3u8" : MimeTypes = "application/vnd.apple.mpegurl" 
      Case "mus" : MimeTypes = "application/vnd.musician" 
      Case "msty" : MimeTypes = "application/vnd.muvee.style" 
      Case "mxml" : MimeTypes = "application/xv+xml" 
      Case "ngdat" : MimeTypes = "application/vnd.nokia.n-gage.data" 
      Case "n-gage" : MimeTypes = "application/vnd.nokia.n-gage.symbian.install" 
      Case "ncx" : MimeTypes = "application/x-dtbncx+xml" 
      Case "nc" : MimeTypes = "application/x-netcdf" 
      Case "nlu" : MimeTypes = "application/vnd.neurolanguage.nlu" 
      Case "dna" : MimeTypes = "application/vnd.dna" 
      Case "nnd" : MimeTypes = "application/vnd.noblenet-directory" 
      Case "nns" : MimeTypes = "application/vnd.noblenet-sealer" 
      Case "nnw" : MimeTypes = "application/vnd.noblenet-web" 
      Case "rpst" : MimeTypes = "application/vnd.nokia.radio-preset" 
      Case "rpss" : MimeTypes = "application/vnd.nokia.radio-presets" 
      Case "n3" : MimeTypes = "text/n3" 
      Case "edm" : MimeTypes = "application/vnd.novadigm.edm" 
      Case "edx" : MimeTypes = "application/vnd.novadigm.edx" 
      Case "ext" : MimeTypes = "application/vnd.novadigm.ext" 
      Case "gph" : MimeTypes = "application/vnd.flographit" 
      Case "ecelp4800" : MimeTypes = "audio/vnd.nuera.ecelp4800" 
      Case "ecelp7470" : MimeTypes = "audio/vnd.nuera.ecelp7470" 
      Case "ecelp9600" : MimeTypes = "audio/vnd.nuera.ecelp9600" 
      Case "oda" : MimeTypes = "application/oda" 
      Case "ogx" : MimeTypes = "application/ogg" 
      Case "oga" : MimeTypes = "audio/ogg" 
      Case "ogv" : MimeTypes = "video/ogg" 
      Case "dd2" : MimeTypes = "application/vnd.oma.dd2+xml" 
      Case "oth" : MimeTypes = "application/vnd.oasis.opendocument.text-web" 
      Case "opf" : MimeTypes = "application/oebps-package+xml" 
      Case "qbo" : MimeTypes = "application/vnd.intu.qbo" 
      Case "oxt" : MimeTypes = "application/vnd.openofficeorg.extension" 
      Case "osf" : MimeTypes = "application/vnd.yamaha.openscoreformat" 
      Case "weba" : MimeTypes = "audio/webm" 
      Case "webm" : MimeTypes = "video/webm" 
      Case "odc" : MimeTypes = "application/vnd.oasis.opendocument.chart" 
      Case "otc" : MimeTypes = "application/vnd.oasis.opendocument.chart-template" 
      Case "odb" : MimeTypes = "application/vnd.oasis.opendocument.database" 
      Case "odf" : MimeTypes = "application/vnd.oasis.opendocument.formula" 
      Case "odft" : MimeTypes = "application/vnd.oasis.opendocument.formula-template" 
      Case "odg" : MimeTypes = "application/vnd.oasis.opendocument.graphics" 
      Case "otg" : MimeTypes = "application/vnd.oasis.opendocument.graphics-template" 
      Case "odi" : MimeTypes = "application/vnd.oasis.opendocument.image" 
      Case "oti" : MimeTypes = "application/vnd.oasis.opendocument.image-template" 
      Case "odp" : MimeTypes = "application/vnd.oasis.opendocument.presentation" 
      Case "otp" : MimeTypes = "application/vnd.oasis.opendocument.presentation-template" 
      Case "ods" : MimeTypes = "application/vnd.oasis.opendocument.spreadsheet" 
      Case "ots" : MimeTypes = "application/vnd.oasis.opendocument.spreadsheet-template" 
      Case "odt" : MimeTypes = "application/vnd.oasis.opendocument.text" 
      Case "odm" : MimeTypes = "application/vnd.oasis.opendocument.text-master" 
      Case "ott" : MimeTypes = "application/vnd.oasis.opendocument.text-template" 
      Case "ktx" : MimeTypes = "image/ktx" 
      Case "sxc" : MimeTypes = "application/vnd.sun.xml.calc" 
      Case "stc" : MimeTypes = "application/vnd.sun.xml.calc.template" 
      Case "sxd" : MimeTypes = "application/vnd.sun.xml.draw" 
      Case "std" : MimeTypes = "application/vnd.sun.xml.draw.template" 
      Case "sxi" : MimeTypes = "application/vnd.sun.xml.impress" 
      Case "sti" : MimeTypes = "application/vnd.sun.xml.impress.template" 
      Case "sxm" : MimeTypes = "application/vnd.sun.xml.math" 
      Case "sxw" : MimeTypes = "application/vnd.sun.xml.writer" 
      Case "sxg" : MimeTypes = "application/vnd.sun.xml.writer.global" 
      Case "stw" : MimeTypes = "application/vnd.sun.xml.writer.template" 
      Case "otf"        : MimeTypes = "application/x-font-otf" 
      Case "osfpvg"     : MimeTypes = "application/vnd.yamaha.openscoreformat.osfpvg+xml" 
      Case "dp"         : MimeTypes = "application/vnd.osgi.dp" 
      Case "pdb"        : MimeTypes = "application/vnd.palm" 
      Case "p"          : MimeTypes = "text/x-pascal" 
      Case "paw"        : MimeTypes = "application/vnd.pawaafile" 
      Case "pclxl"      : MimeTypes = "application/vnd.hp-pclxl" 
      Case "efif"       : MimeTypes = "application/vnd.picsel" 
      Case "pcx"        : MimeTypes = "image/x-pcx" 
      Case "psd"        : MimeTypes = "image/vnd.adobe.photoshop" 
      Case "prf"        : MimeTypes = "application/pics-rules" 
      Case "pic"        : MimeTypes = "image/x-pict" 
      Case "chat"       : MimeTypes = "application/x-chat" 
      Case "p10"        : MimeTypes = "application/pkcs10" 
      Case "p12"        : MimeTypes = "application/x-pkcs12" 
      Case "p7m"        : MimeTypes = "application/pkcs7-mime" 
      Case "p7s"        : MimeTypes = "application/pkcs7-signature" 
      Case "p7r"        : MimeTypes = "application/x-pkcs7-certreqresp" 
      Case "p7b"        : MimeTypes = "application/x-pkcs7-certificates" 
      Case "p8"         : MimeTypes = "application/pkcs8" 
      Case "plf"        : MimeTypes = "application/vnd.pocketlearn" 
      Case "pnm"        : MimeTypes = "image/x-portable-anymap" 
      Case "pbm"        : MimeTypes = "image/x-portable-bitmap" 
      Case "pcf"        : MimeTypes = "application/x-font-pcf" 
      Case "pfr"        : MimeTypes = "application/font-tdpfr" 
      Case "pgn"        : MimeTypes = "application/x-chess-pgn" 
      Case "pgm"        : MimeTypes = "image/x-portable-graymap" 
      Case "png"        : MimeTypes = "image/png" 
      Case "ppm"        : MimeTypes = "image/x-portable-pixmap" 
      Case "pskcxml"    : MimeTypes = "application/pskc+xml" 
      Case "pml"        : MimeTypes = "application/vnd.ctc-posml" 
      Case "ai"         : MimeTypes = "application/postscript" 
      Case "pfa"        : MimeTypes = "application/x-font-type1" 
      Case "pbd"        : MimeTypes = "application/vnd.powerbuilder6" 
      Case "pgp"        : MimeTypes = "application/pgp-encrypted" 
      Case "pgp"        : MimeTypes = "application/pgp-signature" 
      Case "box"        : MimeTypes = "application/vnd.previewsystems.box" 
      Case "ptid"       : MimeTypes = "application/vnd.pvi.ptid1" 
      Case "pls"        : MimeTypes = "application/pls+xml" 
      Case "str"        : MimeTypes = "application/vnd.pg.format" 
      Case "ei6"        : MimeTypes = "application/vnd.pg.osasli" 
      Case "dsc"        : MimeTypes = "text/prs.lines.tag" 
      Case "psf"        : MimeTypes = "application/x-font-linux-psf" 
      Case "qps"        : MimeTypes = "application/vnd.publishare-delta-tree" 
      Case "wg"         : MimeTypes = "application/vnd.pmi.widget" 
      Case "qxd"        : MimeTypes = "application/vnd.quark.quarkxpress" 
      Case "esf"        : MimeTypes = "application/vnd.epson.esf" 
      Case "msf"        : MimeTypes = "application/vnd.epson.msf" 
      Case "ssf"        : MimeTypes = "application/vnd.epson.ssf" 
      Case "qam"        : MimeTypes = "application/vnd.epson.quickanime" 
      Case "qfx"        : MimeTypes = "application/vnd.intu.qfx" 
      Case "qt"         : MimeTypes = "video/quicktime" 
      Case "rar"        : MimeTypes = "application/x-rar-compressed" 
      Case "ram"        : MimeTypes = "audio/x-pn-realaudio" 
      Case "rmp"        : MimeTypes = "audio/x-pn-realaudio-plugin" 
      Case "rsd"        : MimeTypes = "application/rsd+xml" 
      Case "rm"         : MimeTypes = "application/vnd.rn-realmedia" 
      Case "bed"        : MimeTypes = "application/vnd.realvnc.bed" 
      Case "mxl"        : MimeTypes = "application/vnd.recordare.musicxml" 
      Case "musicxml"   : MimeTypes = "application/vnd.recordare.musicxml+xml" 
      Case "rnc"        : MimeTypes = "application/relax-ng-compact-syntax" 
      Case "rdz"        : MimeTypes = "application/vnd.data-vision.rdz" 
      Case "rdf"        : MimeTypes = "application/rdf+xml" 
      Case "rp9"        : MimeTypes = "application/vnd.cloanto.rp9" 
      Case "jisp"       : MimeTypes = "application/vnd.jisp" 
      Case "rtf"        : MimeTypes = "application/rtf" 
      Case "rtx"        : MimeTypes = "text/richtext" 
      Case "link66"     : MimeTypes = "application/vnd.route66.link66+xml" 
      Case "rss"        : MimeTypes = "application/rss+xml" 
      Case "xml"        : MimeTypes = "application/rss+xml" 
      Case "shf"        : MimeTypes = "application/shf+xml" 
      Case "st"         : MimeTypes = "application/vnd.sailingtracker.track" 
      Case "svg"        : MimeTypes = "image/svg+xml" 
      Case "sus"        : MimeTypes = "application/vnd.sus-calendar" 
      Case "sru"        : MimeTypes = "application/sru+xml" 
      Case "setpay"     : MimeTypes = "application/set-payment-initiation" 
      Case "setreg"     : MimeTypes = "application/set-registration-initiation" 
      Case "sema"       : MimeTypes = "application/vnd.sema" 
      Case "semd"       : MimeTypes = "application/vnd.semd" 
      Case "semf"       : MimeTypes = "application/vnd.semf" 
      Case "see"        : MimeTypes = "application/vnd.seemail" 
      Case "snf"        : MimeTypes = "application/x-font-snf" 
      Case "spq"        : MimeTypes = "application/scvp-vp-request" 
      Case "spp"        : MimeTypes = "application/scvp-vp-response" 
      Case "scq"        : MimeTypes = "application/scvp-cv-request" 
      Case "scs"        : MimeTypes = "application/scvp-cv-response" 
      Case "sdp"        : MimeTypes = "application/sdp" 
      Case "etx"        : MimeTypes = "text/x-setext" 
      Case "movie"      : MimeTypes = "video/x-sgi-movie" 
      Case "ifm"        : MimeTypes = "application/vnd.shana.informed.formdata" 
      Case "itp"        : MimeTypes = "application/vnd.shana.informed.formtemplate" 
      Case "iif"        : MimeTypes = "application/vnd.shana.informed.interchange" 
      Case "ipk"        : MimeTypes = "application/vnd.shana.informed.package" 
      Case "tfi"        : MimeTypes = "application/thraud+xml" 
      Case "shar"       : MimeTypes = "application/x-shar" 
      Case "rgb"        : MimeTypes = "image/x-rgb" 
      Case "slt"        : MimeTypes = "application/vnd.epson.salt" 
      Case "aso"        : MimeTypes = "application/vnd.accpac.simply.aso" 
      Case "imp"        : MimeTypes = "application/vnd.accpac.simply.imp" 
      Case "twd"        : MimeTypes = "application/vnd.simtech-mindmapper" 
      Case "csp"        : MimeTypes = "application/vnd.commonspace" 
      Case "saf"        : MimeTypes = "application/vnd.yamaha.smaf-audio" 
      Case "mmf"        : MimeTypes = "application/vnd.smaf" 
      Case "spf"        : MimeTypes = "application/vnd.yamaha.smaf-phrase" 
      Case "teacher"    : MimeTypes = "application/vnd.smart.teacher" 
      Case "svd"        : MimeTypes = "application/vnd.svd" 
      Case "rq"         : MimeTypes = "application/sparql-query" 
      Case "srx"        : MimeTypes = "application/sparql-results+xml" 
      Case "gram"       : MimeTypes = "application/srgs" 
      Case "grxml"      : MimeTypes = "application/srgs+xml" 
      Case "ssml"       : MimeTypes = "application/ssml+xml" 
      Case "skp"        : MimeTypes = "application/vnd.koan" 
      Case "sgml"       : MimeTypes = "text/sgml" 
      Case "sdc"        : MimeTypes = "application/vnd.stardivision.calc" 
      Case "sda"        : MimeTypes = "application/vnd.stardivision.draw" 
      Case "sdd"        : MimeTypes = "application/vnd.stardivision.impress" 
      Case "smf"        : MimeTypes = "application/vnd.stardivision.math" 
      Case "sdw"        : MimeTypes = "application/vnd.stardivision.writer" 
      Case "sgl"        : MimeTypes = "application/vnd.stardivision.writer-global" 
      Case "sm"         : MimeTypes = "application/vnd.stepmania.stepchart" 
      Case "sit"        : MimeTypes = "application/x-stuffit" 
      Case "sitx"       : MimeTypes = "application/x-stuffitx" 
      Case "sdkm"       : MimeTypes = "application/vnd.solent.sdkm+xml" 
      Case "xo"         : MimeTypes = "application/vnd.olpc-sugar" 
      Case "au"         : MimeTypes = "audio/basic" 
      Case "wqd"        : MimeTypes = "application/vnd.wqd" 
      Case "sis"        : MimeTypes = "application/vnd.symbian.install" 
      Case "smi"        : MimeTypes = "application/smil+xml" 
      Case "xsm"        : MimeTypes = "application/vnd.syncml+xml" 
      Case "bdm"        : MimeTypes = "application/vnd.syncml.dm+wbxml" 
      Case "xdm"        : MimeTypes = "application/vnd.syncml.dm+xml" 
      Case "sv4cpio"    : MimeTypes = "application/x-sv4cpio" 
      Case "sv4crc" : MimeTypes = "application/x-sv4crc" 
      Case "sbml" : MimeTypes = "application/sbml+xml" 
      Case "tsv" : MimeTypes = "text/tab-separated-values" 
      Case "tiff" : MimeTypes = "image/tiff" 
      Case "tao" : MimeTypes = "application/vnd.tao.intent-module-archive" 
      Case "tar" : MimeTypes = "application/x-tar" 
      Case "tcl" : MimeTypes = "application/x-tcl" 
      Case "tex" : MimeTypes = "application/x-tex" 
      Case "tfm" : MimeTypes = "application/x-tex-tfm" 
      Case "tei" : MimeTypes = "application/tei+xml" 
      Case "txt" : MimeTypes = "text/plain" 
      Case "dxp" : MimeTypes = "application/vnd.spotfire.dxp" 
      Case "sfs" : MimeTypes = "application/vnd.spotfire.sfs" 
      Case "tsd" : MimeTypes = "application/timestamped-data" 
      Case "tpt" : MimeTypes = "application/vnd.trid.tpt" 
      Case "mxs" : MimeTypes = "application/vnd.triscape.mxs" 
      Case "t" : MimeTypes = "text/troff" 
      Case "tra" : MimeTypes = "application/vnd.trueapp" 
      Case "ttf" : MimeTypes = "application/x-font-ttf" 
      Case "ttl" : MimeTypes = "text/turtle" 
      Case "umj" : MimeTypes = "application/vnd.umajin" 
      Case "uoml" : MimeTypes = "application/vnd.uoml+xml" 
      Case "unityweb" : MimeTypes = "application/vnd.unity" 
      Case "ufd" : MimeTypes = "application/vnd.ufdl" 
      Case "uri" : MimeTypes = "text/uri-list" 
      Case "utz" : MimeTypes = "application/vnd.uiq.theme" 
      Case "ustar" : MimeTypes = "application/x-ustar" 
      Case "uu" : MimeTypes = "text/x-uuencode" 
      Case "vcs" : MimeTypes = "text/x-vcalendar" 
      Case "vcf" : MimeTypes = "text/x-vcard" 
      Case "vcd" : MimeTypes = "application/x-cdlink" 
      Case "vsf" : MimeTypes = "application/vnd.vsf" 
      Case "wrl" : MimeTypes = "model/vrml" 
      Case "vcx" : MimeTypes = "application/vnd.vcx" 
      Case "mts" : MimeTypes = "model/vnd.mts" 
      Case "vtu" : MimeTypes = "model/vnd.vtu" 
      Case "vis" : MimeTypes = "application/vnd.visionary" 
      Case "viv" : MimeTypes = "video/vnd.vivo" 
      Case "ccxml" : MimeTypes = "application/ccxml+xml," 
      Case "vxml" : MimeTypes = "application/voicexml+xml" 
      Case "src" : MimeTypes = "application/x-wais-source" 
      Case "wbxml" : MimeTypes = "application/vnd.wap.wbxml" 
      Case "wbmp" : MimeTypes = "image/vnd.wap.wbmp" 
      Case "wav" : MimeTypes = "audio/x-wav" 
      Case "davmount" : MimeTypes = "application/davmount+xml" 
      Case "woff" : MimeTypes = "application/x-font-woff" 
      Case "wspolicy" : MimeTypes = "application/wspolicy+xml" 
      Case "webp" : MimeTypes = "image/webp" 
      Case "wtb" : MimeTypes = "application/vnd.webturbo" 
      Case "wgt" : MimeTypes = "application/widget" 
      Case "hlp" : MimeTypes = "application/winhlp" 
      Case "wml" : MimeTypes = "text/vnd.wap.wml" 
      Case "wmls" : MimeTypes = "text/vnd.wap.wmlscript" 
      Case "wmlsc" : MimeTypes = "application/vnd.wap.wmlscriptc" 
      Case "wpd" : MimeTypes = "application/vnd.wordperfect" 
      Case "stf" : MimeTypes = "application/vnd.wt.stf" 
      Case "wsdl" : MimeTypes = "application/wsdl+xml" 
      Case "xbm" : MimeTypes = "image/x-xbitmap" 
      Case "xpm" : MimeTypes = "image/x-xpixmap" 
      Case "xwd" : MimeTypes = "image/x-xwindowdump" 
      Case "der" : MimeTypes = "application/x-x509-ca-cert" 
      Case "fig" : MimeTypes = "application/x-xfig" 
      Case "xhtml" : MimeTypes = "application/xhtml+xml" 
      Case "xml" : MimeTypes = "application/xml" 
      Case "xdf" : MimeTypes = "application/xcap-diff+xml" 
      Case "xenc" : MimeTypes = "application/xenc+xml" 
      Case "xer" : MimeTypes = "application/patch-ops-error+xml" 
      Case "rl" : MimeTypes = "application/resource-lists+xml" 
      Case "rs" : MimeTypes = "application/rls-services+xml" 
      Case "rld" : MimeTypes = "application/resource-lists-diff+xml" 
      Case "xslt" : MimeTypes = "application/xslt+xml" 
      Case "xop" : MimeTypes = "application/xop+xml" 
      Case "xpi" : MimeTypes = "application/x-xpinstall" 
      Case "xspf" : MimeTypes = "application/xspf+xml" 
      Case "xul" : MimeTypes = "application/vnd.mozilla.xul+xml" 
      Case "xyz" : MimeTypes = "chemical/x-xyz" 
      Case "yaml" : MimeTypes = "text/yaml" 
      Case "yang" : MimeTypes = "application/yang" 
      Case "yin" : MimeTypes = "application/yin+xml" 
      Case "zir" : MimeTypes = "application/vnd.zul" 
      Case "zip" : MimeTypes = "application/x-zip-compressed" 
      Case "zmm" : MimeTypes = "application/vnd.handheld-entertainment+xml" 
      Case "zaz" : MimeTypes = "application/vnd.zzazz.deck+xml" 
      Case Else       : MimeTypes = "text/plain"
    End Select
  End Function

  '--------------------------------
  ' Get File Name
  '--------------------------------
  Private Function GetFileName(vFile)
    If Instr(1, vFile, "/") <> 0 Then 
      tmpGetFileName = Split(vFile, "/")
      GetFileName = tmpGetFileName( UBound(tmpGetFileName) )
    Else
      GetFileName = vFile
    End If
  End Function 

  '--------------------------------
  ' Get File Name
  '--------------------------------
  Private Function GetFileExtention(vFileExtention)
    If Instr(1, vFileExtention, ".") <> 0 Then 
      ExtentionData = Split(vFileExtention, ".")
      GetFileExtention = ExtentionData( UBound(ExtentionData) )
    Else 
      GetFileExtention = "NOEXTENTION"
    End If
  End Function
  '--------------------------------
  ' Post Action Return: HTTP Status Code
  '--------------------------------
  Public Property Get HTTPStatus()
    HTTPStatus = PostHTTPStatus
  End Property

  '--------------------------------
  ' Post Action Return: HTTP Text
  '--------------------------------
  Public Property Get HTTPAnswer()
    HTTPAnswer = PostHTTPAnswer
  End Property

  '--------------------------------
  ' Form Post URL
  '--------------------------------
  Public Property Get PostURL(vUri)
    letPostURL = vUri
  End Property

  '--------------------------------
  ' Form Type Controller
  '--------------------------------
  Public Property Get FormType(vData)
    If Len(vData) < 3 Then vData = "POST" ' Default Type
    letFormType = vData
  End Property

  '--------------------------------
  ' Form Type Controller
  '--------------------------------
  Public Property Get DebugBoundry()
    DebugBoundry = strDebugBoundry
  End Property

  '--------------------------------
  ' Post The File(s)
  '--------------------------------
  Public Function PostFiles()
    Dim BoundryKey, BoundryStartKey, BoundryEndKey
        BoundryKey = RandomBoundryGenerator()
    Dim BoundryFile()
    Dim BoundryFileSize
        BoundryFileSize = -1

    If IsNull(letPostURL) Then 
        Call ErrorRaise(4, PostToURL)
    End If

    ' Start Boundry
    BoundryStartKey = "------WebKitFormBoundaryd"&BoundryKey&"" & vbCrlf '&_
    '"Content-Disposition: form-data; name=""burak""" & vbcrlf & vbcrlf & "5279" & vbcrlf & "" & _
    '"------WebKitFormBoundaryd"&BoundryKey&"" & vbCrlf &_
    '"Content-Disposition: form-data; name=""dursun""" & vbcrlf & vbcrlf & "0" & vbcrlf
    Boundry         = ""
    ' End Boundry
    BoundryEndKey   = "------WebKitFormBoundaryd"&BoundryKey&"--"

    ' INPUTS
    TotalTextBound = UBound(FileTextArray)
    For i=0 To TotalTextBound
      inputdata = Split( FileTextArray(i), "=|*|=")
      inputname = inputdata(0)
      inputvalue= inputdata(1)

      'If TotalTextBound < i Then 
      'End If
      Boundry=Boundry & "Content-Disposition: form-data; name="""&Trim(inputname)&"""" & vbcrlf&vbcrlf & Trim(inputvalue) & vbcrlf
      If Not i=TotalTextBound Then Boundry=Boundry & "------WebKitFormBoundaryd"&BoundryKey&"" & vbCrlf
    Next

    ' Files
    For i=0 To UBound(FilePostArray)
      If Len(FilePostArray(i)) > 2 Then
        strFileName       = GetFileName( FilePostArray(i) )
        strFileExtention  = GetFileExtention( strFileName )

        strBoundry = "------WebKitFormBoundaryd"&BoundryKey&"" & vbCrlf
        strBoundry=strBoundry&"Content-Disposition: form-data; name=""file"&i&"""; filename=""" & strFileName & """" & vbcrlf & ""
        ' Zip Dosyası İçin
        If MimeTypes( strFileName ) = "application/x-zip-compressed" Then 
          strBoundry=strBoundry&"Content-Transfer-Encoding: base64" & vbcrlf & ""
        End If
        strBoundry=strBoundry&"Content-Type: "& MimeTypes( strFileName ) &"" & vbCrlf & vbCrlf 


        ' Collect
        fBoundry=fBoundry&strBoundry

        ReDim PRESERVE BoundryFile( BoundryFileSize )
        BoundryFileSize=BoundryFileSize+1

        ReDim PRESERVE BoundryFile( BoundryFileSize )
        BoundryFile(BoundryFileSize) = strBoundry
      End If
    Next

    strDebugBoundry=BoundryStartKey&Boundry&fBoundry&BoundryEndKey

    'Response.Write strDebugBoundry
    'Response.End

    Set MainStream = Server.CreateObject("ADODB.Stream")
        MainStream.Type = 1
        MainStream.Mode = 3
        MainStream.Open

        MainStream.Write Stream_StringToBinary( BoundryStartKey , "")
        MainStream.Write Stream_StringToBinary( Boundry , "")
        For i=0 To UBound(FilePostArray)
          strFileName = FilePostArray(i)
          If Len(strFileName) > 2 Then
            strFileExtention  = GetFileExtention( strFileName )

            'Response.Write "<hr>"& BoundryFile(i) &"<hr>"
            MainStream.Write Stream_StringToBinary( BoundryFile(i) , "")
            MainStream.Write ReadBinaryFile( Server.MapPath( strFileName ) )

            If strFileExtention = "zip" OR strFileExtention = "rar" Then 
              MainStream.Write Stream_StringToBinary("00", "")
              MainStream.Write Stream_StringToBinary(""&vbcrlf&"00", "")
            ElseIf strFileExtention = "txt" Then 
              MainStream.Write Stream_StringToBinary("00", "")
            End If
          End If
        Next 
        MainStream.Write Stream_StringToBinary(BoundryEndKey, "")

        MainStream.Position = 0

    Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
        httpRequest.Open letFormType, letPostURL, False, "", "" 
        'httpRequest.SetRequestHeader "Accept", "application/json"
        httpRequest.setRequestHeader "Content-Type", "multipart/form-data; boundary=----WebKitFormBoundaryd"&BoundryKey&""
        httpRequest.Send MainStream.read

    MainStream.Close

    PostHTTPStatus = httpRequest.status
    PostHTTPAnswer = httpRequest.responseText
    If httpRequest.status = 200 Then 
      PostFiles = True
    Else
      PostFiles = False
    End If

    Set httpRequest = Nothing   
    Set MainStream = Nothing  
  End Function
End Class
%>
