using Arcen.HotM.Core;
using Arcen.Universal;

namespace UkrVoice
{
    /// <summary>
    /// Точка входу мода: гра викликає її після завантаження всіх таблиць.
    /// Реєстрація — XMLMods/UkrVoice/ExternalDllInitialLoadCall/UkrVoice_LoadCall.xml.
    /// </summary>
    public class InitialSetupForDLL : IArcenExternalDllInitialLoadCall
    {
        private const string MOD_ID = "UkrVoice";
        private static bool patched = false;

        public void RunAfterAllTableImportsComplete( ArcenExternalDllInitialLoadCall Loader )
        {
            Speaker.Init();
            if ( patched )
                return; //повторне завантаження XML не повинно патчити вдруге
            patched = true;
            int n = HarmonyIntegration.ApplyPatchesFromAssembly( typeof( InitialSetupForDLL ).Assembly, MOD_ID );
            Speaker.Log( "Harmony: пропатчено методів: " + n );
        }

        public void RunImmediatelyOnHandlerProcessed( ArcenExternalDllInitialLoadCall Loader )
        {
        }

        public void RunOnFirstTimeExternalAssemblyLoaded()
        {
        }
    }
}
