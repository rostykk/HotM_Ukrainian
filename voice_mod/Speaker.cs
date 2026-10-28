using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Arcen.Universal;

namespace UkrVoice
{
    /// <summary>
    /// Прототип: пише в UkrVoice_log.txt, що саме гра показала (канал, ключ, український текст),
    /// і передає ключ помічнику speak_helper.ps1, який читає англійський оригінал голосом Windows.
    /// Згодом помічник заміниться програвачем готових аудіофайлів за тим самим ключем.
    /// </summary>
    public static class Speaker
    {
        public const int CH_EVENT = 0, CH_REWARD = 1, CH_HANDBOOK = 2, CH_MINOR = 3, CH_COUNT = 4;
        private static readonly string[] channelNames = { "event", "reward", "handbook", "minor" };
        private static readonly string[] lastToken = new string[CH_COUNT];

        private static bool inited = false;
        private static string modDir = string.Empty;
        private static StreamWriter logWriter = null;
        private static Process helper = null;
        private static int helperFailures = 0;
        private static readonly object sync = new object();

        public static void Init()
        {
            if ( inited )
                return;
            inited = true;
            try
            {
                string gameData = Engine_Universal.OriginalNonTotalConversionGameDataDirectory ?? string.Empty;
                string gameRoot = Path.GetDirectoryName( gameData.TrimEnd( '/', '\\' ) );
                modDir = Path.Combine( Path.Combine( gameRoot, "XMLMods" ), "UkrVoice" );
                logWriter = new StreamWriter( Path.Combine( modDir, "UkrVoice_log.txt" ), true, new UTF8Encoding( false ) );
                logWriter.AutoFlush = true;
                Log( "==== старт гри, тека мода: " + modDir );
            }
            catch ( Exception e )
            {
                ArcenDebugging.LogSingleLine( "UkrVoice: не вдалося відкрити лог: " + e, Verbosity.ShowAsError );
            }
            StartHelper();
        }

        public static void Log( string line )
        {
            lock ( sync )
            {
                try
                {
                    logWriter?.WriteLine( DateTime.Now.ToString( "HH:mm:ss" ) + "\t" + line );
                }
                catch { }
            }
        }

        public static void Error( string where, Exception e )
        {
            Log( "ПОМИЛКА в " + where + ": " + e );
        }

        /// <summary>
        /// Викликається щокадру з патчів; озвучує лише коли вміст вікна змінився.
        /// Помічнику йде рядок SAY\tстиль\tключ\tтекст-у-base64 (UTF-8), щоб у stdin були лише ASCII-символи.
        /// </summary>
        public static void Seen( int channel, string token, string recordKey, string ukText, string voice = "narr" )
        {
            if ( token == lastToken[channel] )
                return;
            lastToken[channel] = token;
            string text = ukText ?? string.Empty;
            string flat = text.Replace( "\r", "" ).Replace( "\n", " / " );
            if ( flat.Length > 300 )
                flat = flat.Substring( 0, 300 ) + "…";
            Log( channelNames[channel] + "\t" + recordKey + "\t" + flat );
            lastMainKey = recordKey;
            lastMainText = text;
            Send( "SAY\t" + voice + "\t" + recordKey + "\t" + Convert.ToBase64String( Encoding.UTF8.GetBytes( text ) ) + "\tmain" );
        }

        //останній основний текст вікна — для повтору по F9
        private static string lastMainKey = null, lastMainText = null;

        #region Підказки до варіантів: самі після затримки, Ctrl — стоп / ще раз
        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern short GetAsyncKeyState( int vKey );
        private const int VK_CONTROL = 0x11;
        private const int VK_F9 = 0x78; //F9 — повторити основний текст (у грі не зайнята; F10 зайнята)
        private static bool f9WasDown = false;
        private const float HOVER_DWELL_SECONDS = 0.5f;  //стільки миша має простояти на варіанті, щоб почати читати
        private const float HOVER_FRESH_SECONDS = 0.25f; //підказку перемальовують щокадру; довше без неї — вона зникла

        private static bool ctrlWasDown = false;
        private static string hoverKey = null, hoverText = null;
        private static float hoverStart = -100f, hoverTime = -100f;
        private static bool hoverDone = false;          //цю підказку вже озвучено (або зупинено Ctrl) — сама вдруге не почне
        private static string speakingHoverKey = null;  //яку підказку зараз читаємо

        /// <summary>Гра щойно намалювала підказку до варіанта.</summary>
        public static void Hover( string key, string text )
        {
            float now = UnityEngine.Time.unscaledTime;
            if ( key != hoverKey || now - hoverTime > HOVER_FRESH_SECONDS ) //нова підказка або повернулися після паузи
            {
                hoverKey = key;
                hoverStart = now;
                hoverDone = false;
            }
            hoverText = text;
            hoverTime = now;
            Poll();
        }

        /// <summary>Як Hover, але текст складаємо лише для нової підказки (дорого збирати його щокадру).</summary>
        public static void HoverLazy( string key, Func<string> makeText )
        {
            if ( key == hoverKey && hoverText != null && UnityEngine.Time.unscaledTime - hoverTime <= HOVER_FRESH_SECONDS )
                Hover( key, hoverText );
            else
                Hover( key, makeText() ?? string.Empty );
        }

        /// <summary>Викликається щокадру з патчів: затримка, зникнення підказки, Ctrl.</summary>
        public static void Poll()
        {
            float now = UnityEngine.Time.unscaledTime;
            bool fresh = hoverText != null && now - hoverTime <= HOVER_FRESH_SECONDS;

            if ( speakingHoverKey != null && !fresh ) //підказка зникла — замовкає лише підказка, не основний текст
            {
                speakingHoverKey = null;
                Log( "hover\t(підказка зникла)" );
                Send( "STOP_HOVER" );
            }
            else if ( fresh && !hoverDone && now - hoverStart >= HOVER_DWELL_SECONDS )
                SpeakHover( "затримка" );

            bool down, f9;
            try
            {
                down = ( GetAsyncKeyState( VK_CONTROL ) & 0x8000 ) != 0;
                f9 = ( GetAsyncKeyState( VK_F9 ) & 0x8000 ) != 0;
            }
            catch { return; }
            if ( down && !ctrlWasDown && fresh )
            {
                if ( speakingHoverKey == hoverKey )
                {
                    speakingHoverKey = null;
                    hoverDone = true;
                    Log( "hover\t(стоп по Ctrl)" );
                    Send( "STOP_HOVER" );
                }
                else
                    SpeakHover( "Ctrl" );
            }
            ctrlWasDown = down;

            if ( f9 && !f9WasDown && lastMainText != null )
            {
                speakingHoverKey = null;
                Log( "F9\tповтор " + lastMainKey );
                Send( "SAY\tnarr\t" + lastMainKey + "\t" + Convert.ToBase64String( Encoding.UTF8.GetBytes( lastMainText ) ) + "\tmain" );
            }
            f9WasDown = f9;
        }

        private static void SpeakHover( string why )
        {
            hoverDone = true;
            speakingHoverKey = hoverKey;
            string flat = hoverText.Replace( "\r", "" ).Replace( "\n", " / " );
            Log( "hover(" + why + ")\t" + hoverKey + "\t" + ( flat.Length > 300 ? flat.Substring( 0, 300 ) + "…" : flat ) );
            //hover_auto помічник пропустить, якщо зараз звучить основний текст; hover (Ctrl) — перебиває
            Send( "SAY\tnarr\t" + hoverKey + "\t" + Convert.ToBase64String( Encoding.UTF8.GetBytes( hoverText ) )
                + ( why == "Ctrl" ? "\thover" : "\thover_auto" ) );
        }
        #endregion

        /// <summary>Склеює непорожні частини тексту через перенос рядка.</summary>
        public static string Join( params string[] parts )
        {
            StringBuilder sb = new StringBuilder();
            foreach ( string p in parts )
            {
                if ( string.IsNullOrEmpty( p ) || p.Trim().Length == 0 )
                    continue;
                if ( sb.Length > 0 )
                    sb.Append( '\n' );
                sb.Append( p );
            }
            return sb.ToString();
        }

        /// <summary>Вікно закрилося: забуваємо останній вміст і обриваємо голос.</summary>
        public static void Closed( int channel )
        {
            if ( lastToken[channel] == null )
                return;
            lastToken[channel] = null;
            Log( channelNames[channel] + "\t(закрито)" );
            Send( "STOP" );
        }

        private static void StartHelper()
        {
            if ( helperFailures > 3 || string.IsNullOrEmpty( modDir ) )
                return;
            try
            {
                //helper_cmd.txt (необов'язковий): 1-й рядок — програма, 2-й — аргументи; {mod} = тека мода.
                //Без нього — голос Windows (speak_helper.ps1, англійський оригінал).
                string exe = "powershell.exe";
                string arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + Path.Combine( modDir, "speak_helper.ps1" ) + "\"";
                string cfg = Path.Combine( modDir, "helper_cmd.txt" );
                if ( File.Exists( cfg ) )
                {
                    string[] lines = File.ReadAllLines( cfg, Encoding.UTF8 );
                    if ( lines.Length >= 2 && lines[0].Trim().Length > 0 )
                    {
                        exe = lines[0].Trim().Replace( "{mod}", modDir );
                        arguments = lines[1].Trim().Replace( "{mod}", modDir );
                    }
                }
                ProcessStartInfo psi = new ProcessStartInfo( exe, arguments );
                psi.UseShellExecute = false;
                psi.RedirectStandardInput = true;
                psi.CreateNoWindow = true;
                psi.WorkingDirectory = modDir;
                helper = Process.Start( psi );
                Log( "помічник озвучки запущено: " + exe + " " + arguments + ", pid " + helper?.Id );
            }
            catch ( Exception e )
            {
                helperFailures++;
                helper = null;
                Error( "StartHelper", e );
            }
        }

        private static void Send( string line )
        {
            try
            {
                if ( helper == null || helper.HasExited )
                    StartHelper();
                if ( helper == null )
                    return;
                helper.StandardInput.WriteLine( line ); //ключі ASCII, кодування не важливе
                helper.StandardInput.Flush();
            }
            catch ( Exception e )
            {
                helperFailures++;
                helper = null;
                Error( "Send", e );
            }
        }
    }
}
