#region ensamblado SolidEdge.Community, Version=108.0.0.0, Culture=neutral, PublicKeyToken=null
// C:\Users\elfst\OneDrive\Documentos\Visual_Studio_2022\CS\Print_Bom\packages\SolidEdge.Community.108.2.0\lib\net40\SolidEdge.Community.dll
// Decompiled with ICSharpCode.Decompiler 8.1.1.7464
#endregion

using System;

namespace SolidEdgeCommunity
{
    //
    // Resumen:
    //     Generic class used to execute an IsolatedTaskProxy implementation.
    //
    // Parámetros de tipo:
    //   T:
    public sealed class IsolatedTask<T> : IDisposable where T : IsolatedTaskProxy
    {
        private Type _proxyType;

        private AppDomain _appDomain;

        private T _proxy;

        public T Proxy => _proxy;

        public IsolatedTask()
        {
            _proxyType = typeof(T);
            _appDomain = AppDomain.CreateDomain($"{_proxyType.Name} AppDomain", null, AppDomain.CurrentDomain.SetupInformation);
            _proxy = (T)_appDomain.CreateInstanceAndUnwrap(_proxyType.Assembly.FullName, _proxyType.FullName);
        }

        void IDisposable.Dispose()
        {
            if (_appDomain != null)
            {
                AppDomain.Unload(_appDomain);
            }

            _proxy = null;
            _appDomain = null;
            _proxyType = null;
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
    Cargar desde: "C:\Users\elfst\OneDrive\Documentos\Visual_Studio_2022\CS\Print_Bom\packages\Interop.SolidEdge.108.4.0\lib\net40\Interop.SolidEdge.dll"
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
}
