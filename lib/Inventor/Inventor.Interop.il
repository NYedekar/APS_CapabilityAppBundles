// Inventor.Interop.il
// Minimal interop assembly declaring Inventor.ApplicationAddInServer.
//
// This IL assembly exists so InventorBOMExtractor.csproj can reference it with
// EmbedInteropTypes=True. The embedded type has tdComImport (set by the 'import'
// keyword here), which is required for:
//   1. [TypeIdentifier] to pass the type-equivalence validation
//   2. The CLR CCW to expose IApplicationAddInServer via QueryInterface
//
// Without a separate assembly + EmbedInteropTypes, [ComImport] declared in the
// same assembly is NOT exposed by the CCW, causing E_NOINTERFACE from Inventor.
//
// All attribute blobs are copy-exact from UpdateIPTParam.dll (official Autodesk
// DA sample) TypeDef[4] Inventor.ApplicationAddInServer.
//
// Compile on Windows (CI):
//   ilasm.exe /dll /quiet /output:Inventor.Interop.dll Inventor.Interop.il

.assembly extern mscorlib
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89)
  .ver 4:0:0:0
}

.assembly Inventor.Interop
{
  .ver 1:0:0:0

  // [assembly: Guid("D98A091D-3A0F-4C3E-B36E-61F62068D488")]
  // Required by CS1747: EmbedInteropTypes needs a typelib GUID on the assembly.
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string)
       = ( 01 00 24 44 39 38 41 30 39 31 44 2D 33 41 30 46 2D 34 43 33 45 2D 42 33 36 45 2D 36 31 46 36 32 30 36 38 44 34 38 38 00 00 )

  // [assembly: ImportedFromTypeLib("Inventor")]
  // Required by CS1759: EmbedInteropTypes needs ImportedFromTypeLib or PrimaryInteropAssembly.
  .custom instance void [mscorlib]System.Runtime.InteropServices.ImportedFromTypeLibAttribute::.ctor(string)
       = ( 01 00 08 49 6E 76 65 6E 74 6F 72 00 00 )
}

.module Inventor.Interop.dll

// 'import' sets tdComImport — required for [TypeIdentifier] type-equivalence to work
.class interface public abstract auto ansi import
       Inventor.ApplicationAddInServer
{
  // [Guid("E3571293-DB40-11D2-B783-0060B0F159EF")]
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string)
       = ( 01 00 24 45 33 35 37 31 32 39 33 2D 44 42 34 30 2D 31 31 44 32 2D 42 37 38 33 2D 30 30 36 30 42 30 46 31 35 39 45 46 00 00 )

  // [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]  (value = 2)
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(int16)
       = ( 01 00 02 00 00 00 )

  // void Activate(object addInSiteObject, bool firstTime)  [DispId(0x03001201)]
  // DISPIDs are Inventor's REAL values (extracted from UpdateIPTParam.dll), NOT 1/2/3/4.
  // IApplicationAddInServer is InterfaceIsIDispatch, so InventorCoreConsole calls every
  // member by these exact DISPIDs. Wrong DISPIDs => get_Automation returns
  // DISP_E_MEMBERNOTFOUND => no automation object => Run never reached.
  // ImplFlags=0x1083 (runtime|internalcall|preservesig) — matches UpdateIPTParam.dll exactly.
  .method public hidebysig newslot virtual abstract
          instance void  Activate([in] object addInSiteObject,
                                   [in] bool firstTime) runtime managed internalcall preservesig
  {
    // DispId(0x03001201)  blob: 01 00 <01 12 00 03> 00 00
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32)
         = ( 01 00 01 12 00 03 00 00 )
  }

  // void Deactivate()  [DispId(0x03001202)]
  .method public hidebysig newslot virtual abstract
          instance void  Deactivate() runtime managed internalcall preservesig
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32)
         = ( 01 00 02 12 00 03 00 00 )
  }

  // void ExecuteCommand(int commandID)  [DispId(0x03001203)]
  .method public hidebysig newslot virtual abstract
          instance void  ExecuteCommand([in] int32 commandID) runtime managed internalcall preservesig
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32)
         = ( 01 00 03 12 00 03 00 00 )
  }

  // object get_Automation()  [DispId(0x03001204)]
  // preservesig is critical: without it the CCW treats object return as HRESULT+outparam,
  // discarding the BOMExtractorAutomation instance and returning null to InventorCoreConsole.
  .method public hidebysig newslot specialname virtual abstract
          instance object  get_Automation() runtime managed internalcall preservesig
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.DispIdAttribute::.ctor(int32)
         = ( 01 00 04 12 00 03 00 00 )
  }

  .property instance object Automation()
  {
    .get instance object Inventor.ApplicationAddInServer::get_Automation()
  }
}
