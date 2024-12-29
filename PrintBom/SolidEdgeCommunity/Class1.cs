#region ensamblado SolidEdge.Community, Version=108.0.0.0, Culture=neutral, PublicKeyToken=null
// C:\Users\elfst\OneDrive\Documentos\Visual_Studio_2022\CS\Nueva carpeta\Print_Bom\Print_Bom\packages\SolidEdge.Community.108.2.0\lib\net40\SolidEdge.Community.dll
// Decompiled with ICSharpCode.Decompiler 8.1.1.7464
#endregion

using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace SolidEdgeCommunity;

//
// Resumen:
//     Class that implements the OLE IMessageFilter interface.
public class OleMessageFilter : IMessageFilter
{
    [DllImport("Ole32.dll")]
    private static extern int CoRegisterMessageFilter(IMessageFilter newFilter, out IMessageFilter oldFilter);

    //
    // Resumen:
    //     Private constructor.
    //
    // Comentarios:
    //     Instance of this class is only created by Register().
    private OleMessageFilter()
    {
    }

    //
    // Resumen:
    //     Destructor.
    ~OleMessageFilter()
    {
        Unregister();
    }

    //
    // Resumen:
    //     Registers this instance of IMessageFilter with OLE to handle concurrency issues
    //     on the current thread.
    //
    // Comentarios:
    //     Only one message filter can be registered for each thread. Threads in multithreaded
    //     apartments cannot have message filters. Thread.CurrentThread.GetApartmentState()
    //     must equal ApartmentState.STA. In console applications, this can be achieved
    //     by applying the STAThreadAttribute to the Main() method. In WinForm applications,
    //     it is default.
    public static void Register()
    {
        IMessageFilter newFilter = new OleMessageFilter();
        IMessageFilter oldFilter = null;
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
        {
            Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(newFilter, out oldFilter));
            return;
        }

        throw new Exception("The current thread's apartment state must be STA.");
    }

    //
    // Resumen:
    //     Unregisters a previous instance of IMessageFilter with OLE on the current thread.
    //
    //
    // Comentarios:
    //     It is not necessary to call Unregister() unless you need to explicitly do so
    //     as it is handled in the destructor.
    public static void Unregister()
    {
        IMessageFilter oldFilter = null;
        CoRegisterMessageFilter(null, out oldFilter);
    }

    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
    {
        return 0;
    }

    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
    {
        return 2;
    }

    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == 2)
        {
            return 99;
        }

        return -1;
    }
}
#if false // Registro de descompilación
"18" elementos en caché
------------------
Resolver: "mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Se encontró un solo ensamblado: "mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Cargar desde: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\mscorlib.dll"
------------------
Resolver: "System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Se encontró un solo ensamblado: "System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Cargar desde: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.Core.dll"
------------------
Resolver: "Interop.SolidEdge, Version=108.0.0.0, Culture=neutral, PublicKeyToken=null"
Se encontró un solo ensamblado: "Interop.SolidEdge, Version=108.0.0.0, Culture=neutral, PublicKeyToken=null"
Cargar desde: "C:\Users\elfst\OneDrive\Documentos\Visual_Studio_2022\CS\Nueva carpeta\Print_Bom\Print_Bom\packages\Interop.SolidEdge.108.4.0\lib\net40\Interop.SolidEdge.dll"
------------------
Resolver: "System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
Se encontró un solo ensamblado: "System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
Cargar desde: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.Drawing.dll"
------------------
Resolver: "System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Se encontró un solo ensamblado: "System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Cargar desde: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.Windows.Forms.dll"
------------------
Resolver: "System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Se encontró un solo ensamblado: "System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Cargar desde: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.dll"
#endif
