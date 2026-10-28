using System;
using System.Reflection;
using Arcen.HotM.Core;
using Arcen.Universal;
using Arcen.HotM.ExternalVis;
using HarmonyLib;

namespace UkrVoice
{
    // Усі патчі — Postfix на методах, що щокадру пишуть текст тіла вікна.
    // Самі вони нічого не змінюють у грі; лише дивляться, ЩО зараз показано, і повідомляють Speaker.

    #region Події (NPCEvent): вікно великих подій
    [HarmonyPatch( typeof( Window_MajorEventWindow ), nameof( Window_MajorEventWindow.tBodyText_WriteBody ) )]
    public static class Patch_MajorEvent_Body
    {
        private static readonly FieldInfo extraField = AccessTools.Field( typeof( Window_MajorEventWindow ), "ExtraBodyText" );

        public static void Postfix()
        {
            try
            {
                Speaker.Poll();
                NPCEvent ev = SimCommon.CurrentEvent;
                if ( ev == null )
                    return;
                string extra = extraField?.GetValue( null ) as string;
                if ( !string.IsNullOrEmpty( extra ) ) //після вибору гра показує текст результату
                {
                    Speaker.Seen( Speaker.CH_EVENT, "R:" + ev.ID + ":" + extra.Length, "RESULT|" + ev.ID, extra );
                    return;
                }
                Speaker.Seen( Speaker.CH_EVENT, "E:" + ev.ID, "NPCEvent|" + ev.ID, Speaker.Join( ev.GetDisplayName(), ev.HighText.Text, ev.GetDescription(), ev.StrategyTip.Text ) );
            }
            catch ( Exception e ) { Speaker.Error( "MajorEvent", e ); }
        }
    }

    [HarmonyPatch( typeof( Window_MajorEventWindow ), nameof( Window_MajorEventWindow.Close ) )]
    public static class Patch_MajorEvent_Close
    {
        public static void Postfix() { Speaker.Closed( Speaker.CH_EVENT ); }
    }
    #endregion

    #region Малі події (NPCEvent через MinorEventHandler): вікно простого вибору
    [HarmonyPatch( typeof( Window_SimpleChoiceWindow ), nameof( Window_SimpleChoiceWindow.tBodyText_WriteBody ) )]
    public static class Patch_SimpleChoice_Body
    {
        private static readonly FieldInfo extraField = AccessTools.Field( typeof( Window_SimpleChoiceWindow ), "ExtraBodyText" );

        public static void Postfix()
        {
            try
            {
                Speaker.Poll();
                ISimpleChoiceProvider p = SimCommon.CurrentSimpleChoice;
                if ( p == null )
                    return;
                string extra = extraField?.GetValue( null ) as string;
                if ( !string.IsNullOrEmpty( extra ) ) //текст результату після вибору
                {
                    Speaker.Seen( Speaker.CH_MINOR, "R:" + extra.Length + ":" + extra.GetHashCode(), "RESULT|minor", extra );
                    return;
                }
                if ( p is MinorEventHandler mh && mh.MinorEvent != null )
                    Speaker.Seen( Speaker.CH_MINOR, "E:" + mh.MinorEvent.ID, "NPCEvent|" + mh.MinorEvent.ID, Speaker.Join( mh.MinorEvent.GetDisplayName(), mh.MinorEvent.HighText.Text, mh.MinorEvent.GetDescription(), mh.MinorEvent.StrategyTip.Text ) );
                else
                    Speaker.Seen( Speaker.CH_MINOR, "P:" + p.GetType().FullName, "PROVIDER|" + p.GetType().Name, string.Empty );
            }
            catch ( Exception e ) { Speaker.Error( "SimpleChoice", e ); }
        }
    }

    [HarmonyPatch( typeof( Window_SimpleChoiceWindow ), nameof( Window_SimpleChoiceWindow.OnHideAfterNotShowing ) )]
    public static class Patch_SimpleChoice_Hide
    {
        public static void Postfix() { Speaker.Closed( Speaker.CH_MINOR ); }
    }
    #endregion

    #region Розмови з NPC, ключові повідомлення та інше: вікно нагород
    [HarmonyPatch( typeof( Window_RewardWindow ), nameof( Window_RewardWindow.tBodyText_WriteBody ) )]
    public static class Patch_Reward_Body
    {
        private static readonly FieldInfo extraField = AccessTools.Field( typeof( Window_RewardWindow ), "ExtraBodyText" );

        public static void Postfix()
        {
            try
            {
                Speaker.Poll();
                IRewardOptionProvider p = SimCommon.RewardProvider;
                if ( p == null )
                    return;
                string extra = extraField?.GetValue( null ) as string;
                if ( !string.IsNullOrEmpty( extra ) )
                {
                    Speaker.Seen( Speaker.CH_REWARD, "R:" + extra.Length + ":" + extra.GetHashCode(), "RESULT|" + p.GetType().Name, extra );
                    return;
                }
                if ( p is NPCDialogChoiceHandler dh && dh.Dialog != null )
                    Speaker.Seen( Speaker.CH_REWARD, "D:" + dh.Dialog.ID, "NPCDialog|" + dh.Dialog.ID, Speaker.Join( dh.Dialog.GetDisplayName(), dh.Dialog.HighText.Text, dh.Dialog.GetDescription(), dh.Dialog.StrategyTip.Text ) );
                else if ( p is ProjectOutcomeHandler ph && ph.Project != null ) //вибір цілі/результату проєкту
                {
                    MachineProject pr = ph.Project;
                    bool alt = pr.UseAltTextIfGFlagTrue?.DGD?.IsTripped ?? false;
                    Speaker.Seen( Speaker.CH_REWARD, "J:" + pr.ID + ( alt ? ":alt" : "" ), "MachineProject|" + pr.ID,
                        alt ? Speaker.Join( pr.GetDisplayName(), pr.AltDescription.Text, pr.AltStrategyTip.Text )
                            : Speaker.Join( pr.GetDisplayName(), pr.GetDescription(), pr.StrategyTip.Text ) );
                }
                else if ( p is SimpleMessageRewardHandler mh && mh.Message != null )
                    Speaker.Seen( Speaker.CH_REWARD, "M:" + mh.Message.ID, "OtherKeyMessage|" + mh.Message.ID, Speaker.Join( mh.Message.RewardHeaderText.Text, mh.Message.GetDisplayName(), mh.Message.HighText.Text, mh.Message.GetDescription(), mh.Message.StrategyTip.Text ) );
                else //результати проєктів, натхнення тощо — поки лише фіксуємо тип
                    Speaker.Seen( Speaker.CH_REWARD, "P:" + p.GetType().FullName, "PROVIDER|" + p.GetType().Name, string.Empty );
            }
            catch ( Exception e ) { Speaker.Error( "Reward", e ); }
        }
    }

    [HarmonyPatch( typeof( Window_RewardWindow ), nameof( Window_RewardWindow.OnHideAfterNotShowing ) )]
    public static class Patch_Reward_Hide
    {
        public static void Postfix() { Speaker.Closed( Speaker.CH_REWARD ); }
    }
    #endregion

    #region Підказки до варіантів (озвучення по Ctrl)
    [HarmonyPatch( typeof( EventChoice ), nameof( EventChoice.RenderTooltip ) )]
    public static class Patch_EventChoice_Tooltip
    {
        public static void Postfix( EventChoice __instance, bool __result )
        {
            try
            {
                if ( !__result )
                    return;
                Speaker.Hover( "NPCEvent|" + ( __instance.ParentEventOrNull?.ID ?? "?" ) + "|choice|" + __instance.ID,
                    Speaker.Join( __instance.DisplayName.Text, __instance.Description.Text, __instance.StrategyTip.Text ) );
            }
            catch ( Exception e ) { Speaker.Error( "EventChoiceTooltip", e ); }
        }
    }

    [HarmonyPatch( typeof( NPCDialogChoice ), nameof( NPCDialogChoice.RenderRewardTooltipText ) )]
    public static class Patch_DialogChoice_Tooltip
    {
        public static void Postfix( NPCDialogChoice __instance )
        {
            try
            {
                Speaker.Hover( "NPCDialogChoice|" + __instance.ID,
                    Speaker.Join( __instance.DisplayName.Text, __instance.Description.Text, __instance.StrategyTip.Text ) );
            }
            catch ( Exception e ) { Speaker.Error( "DialogChoiceTooltip", e ); }
        }
    }

    [HarmonyPatch( typeof( ProjectOutcome ), nameof( ProjectOutcome.RenderRewardTooltipText ) )]
    public static class Patch_ProjectOutcome_Tooltip
    {
        public static void Postfix( ProjectOutcome __instance )
        {
            try
            {
                bool alt = __instance.UseAltTextIfGFlagTrue?.DGD?.IsTripped ?? false;
                Speaker.Hover( "MachineProject|" + ( __instance.ParentProject?.ID ?? "?" ) + "|available_outcome|" + __instance.ShortID,
                    alt ? Speaker.Join( __instance.DisplayName.Text, __instance.AltDescription.Text, __instance.AltStrategyTip.Text )
                        : Speaker.Join( __instance.DisplayName.Text, __instance.Description.Text, __instance.StrategyTip.Text ) );
            }
            catch ( Exception e ) { Speaker.Error( "ProjectOutcomeTooltip", e ); }
        }
    }

    [HarmonyPatch( typeof( OtherKeyMessageOption ), nameof( OtherKeyMessageOption.RenderRewardTooltipText ) )]
    public static class Patch_MessageOption_Tooltip
    {
        public static void Postfix( OtherKeyMessageOption __instance )
        {
            try
            {
                Speaker.Hover( "OtherKeyMessageOption|" + __instance.ID,
                    Speaker.Join( __instance.Line1.Text, __instance.Line2.Text, __instance.Description.Text, __instance.StrategyTip.Text ) );
            }
            catch ( Exception e ) { Speaker.Error( "MessageOptionTooltip", e ); }
        }
    }
    #endregion

    #region Список завдань і проєктів справа (підказки при наведенні)
    [HarmonyPatch( typeof( CityTask ), nameof( CityTask.RenderToastTooltipText ),
        new Type[] { typeof( IArcenUIElementForSizing ), typeof( SideClamp ), typeof( TooltipExtraRules ) } )]
    public static class Patch_CityTask_Tooltip
    {
        private static readonly ArcenDoubleCharacterBuffer buf = new ArcenDoubleCharacterBuffer( "UkrVoice-CityTask" );

        private static string Part( CityTask task, CityTaskTextType type )
        {
            buf.EnsureResetForNextUpdate();
            task.Implementation.WriteText( task, type, buf ); //той самий текст, що реалізація пише в підказку
            return buf.GetStringAndResetForNextUpdate();
        }

        public static void Postfix( CityTask __instance )
        {
            try
            {
                if ( __instance?.Implementation == null )
                    return;
                Speaker.HoverLazy( "CityTask|" + __instance.ID, () => Speaker.Join(
                    Part( __instance, CityTaskTextType.DisplayName ),
                    Part( __instance, CityTaskTextType.Description ),
                    Part( __instance, CityTaskTextType.StrategyText ) ) );
            }
            catch ( Exception e ) { Speaker.Error( "CityTaskTooltip", e ); }
        }
    }

    [HarmonyPatch( typeof( MachineProject ), nameof( MachineProject.RenderProjectTooltip ) )]
    public static class Patch_Project_Tooltip
    {
        public static void Postfix( MachineProject __instance )
        {
            try
            {
                bool alt = __instance.UseAltTextIfGFlagTrue?.DGD?.IsTripped ?? false;
                Speaker.HoverLazy( "MachineProject|" + __instance.ID + ( alt ? ":alt" : "" ), () => alt
                    ? Speaker.Join( __instance.GetDisplayName(), __instance.AltDescription.Text, __instance.AltStrategyTip.Text )
                    : Speaker.Join( __instance.GetDisplayName(), __instance.GetDescription(), __instance.StrategyTip.Text ) );
            }
            catch ( Exception e ) { Speaker.Error( "ProjectTooltip", e ); }
        }
    }
    #endregion

    #region Довідник
    [HarmonyPatch( typeof( Window_Handbook ), nameof( Window_Handbook.tLeftMainText_FillTextPerFrame ) )]
    public static class Patch_Handbook_Text
    {
        public static void Postfix()
        {
            try
            {
                MachineHandbookEntry entry = Window_Handbook.EntryToShow;
                if ( entry == null )
                    return;
                Speaker.Seen( Speaker.CH_HANDBOOK, "H:" + entry.ID, "MachineHandbookEntry|" + entry.ID, Speaker.Join( entry.GetDescription(), entry.FinalStrategyTip ) );
            }
            catch ( Exception e ) { Speaker.Error( "Handbook", e ); }
        }
    }

    [HarmonyPatch( typeof( Window_Handbook ), nameof( Window_Handbook.OnClose ) )]
    public static class Patch_Handbook_Close
    {
        public static void Postfix() { Speaker.Closed( Speaker.CH_HANDBOOK ); }
    }
    #endregion
}
