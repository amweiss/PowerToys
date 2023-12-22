// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Diagnostics;
using System.Globalization;
using ManagedCommon;
using Microsoft.PowerToys.Settings.UI.Library;
using Windows.ApplicationModel.Appointments;
using Wox.Plugin;
using Wox.Plugin.Logger;
using BrowserInfo = Wox.Plugin.Common.DefaultBrowserInfo;

namespace Community.PowerToys.Run.Plugin.Agenda
{
    public class Main : IPlugin, IPluginI18n, IDisposable, ISettingProvider
    {
        // Should only be set in Init()
        private Action? onPluginError;

        public string Name => Properties.Resources.plugin_name;

        public string Description => Properties.Resources.plugin_description;

        public static string PluginID => "C3F7E8A14E5B4E6C8E9F0D2A7B9C0D1E";

        private PluginInitContext? Context { get; set; }

        private string? IconPath { get; set; }

        private bool _disposed;

        public void Init(PluginInitContext context)
        {
            Context = context ?? throw new ArgumentNullException(paramName: nameof(context));
            BrowserInfo.UpdateIfTimePassed();

            onPluginError = () =>
            {
                string errorMsgString = string.Format(CultureInfo.CurrentCulture, Properties.Resources.plugin_search_failed, BrowserInfo.Name ?? BrowserInfo.MSEdgeName);

                Log.Error(errorMsgString, this.GetType());
                Context.API.ShowMsg(
                    $"Plugin: {Properties.Resources.plugin_name}",
                    errorMsgString);
            };

            Context.API.ThemeChanged += OnThemeChanged;
            UpdateIconPath(Context.API.GetCurrentTheme());
        }

        public List<Result> Query(Query query)
        {
            ArgumentNullException.ThrowIfNull(query);

            return GetAgendaResultsAsync(query).GetAwaiter().GetResult();
        }

        private async Task<List<Result>> GetAgendaResultsAsync(Query query)
        {
            try
            {
                var user = Windows.System.User.GetDefault();
                var appointmentManager = AppointmentManager.GetForUser(user);
                var appointmentStore = await appointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadOnly);

                var duration = TimeSpan.FromHours(24);
                var appointments = await appointmentStore.FindAppointmentsAsync(DateTime.Now, duration);
                var nextAppointment = appointments?[0];

                if (nextAppointment == null)
                {
                    return new List<Result>();
                }

                return new List<Result>()
                {
                    new Result
                    {
                        Title = "Agenda",
                        SubTitle = @$"Join {nextAppointment.Subject}",
                        ContextData = nextAppointment,
                        ToolTipData = new ToolTipData(Name, @$"{nextAppointment.Subject}@{nextAppointment.StartTime}"),
                        IcoPath = IconPath,
                        Action = _ =>
                        {
                            Process.Start(new ProcessStartInfo(nextAppointment.OnlineMeetingLink));
                            Console.WriteLine(nextAppointment);
                            return true;
                        },
                    },
                };
            }
            catch (Exception)
            {
                // Any other crash occurred
                // We want to keep the process alive if any the mages library throws any exceptions.
            }

            var results = new List<Result>();
            return results;
        }

        public System.Windows.Controls.Control CreateSettingPanel()
        {
            return new System.Windows.Controls.Control();
        }

        public IEnumerable<PluginAdditionalOption> AdditionalOptions
        {
            get
            {
                return new List<PluginAdditionalOption>();
            }
        }

        public void UpdateSettings(PowerLauncherPluginSettings settings)
        {
        }

        public string GetTranslatedPluginTitle()
        {
            return Properties.Resources.plugin_name;
        }

        public string GetTranslatedPluginDescription()
        {
            return Properties.Resources.plugin_description;
        }

        private void OnThemeChanged(Theme currentTheme, Theme newTheme)
        {
            UpdateIconPath(newTheme);
        }

        private void UpdateIconPath(Theme theme) =>
            IconPath =
                theme == Theme.Light || theme == Theme.HighContrastWhite
            ? "Images/agenda.light.png"
            : "Images/agenda.dark.png";

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (Context != null && Context.API != null)
                    {
                        Context.API.ThemeChanged -= OnThemeChanged;
                    }

                    _disposed = true;
                }
            }
        }
    }
}
