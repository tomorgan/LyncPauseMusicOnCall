using Microsoft.Lync.Model;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PauseMusicOnCall
{
    public class VM : INotifyPropertyChanged
    {
        private Client _client;
        private bool _pauseEnabled;
        public event PropertyChangedEventHandler PropertyChanged;

        //Properties for the view
        public string SelfSipAddress { get; private set; }
        public string PauseStatus { get; private set; }

        // Stuff needed to send the Play/Pause key
        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        public const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        public const int KEYEVENTF_KEYUP = 0x0002; //Key up flag
        public const int VK_LCONTROL = 0xA2; //Left Control key code
        public const int A = 0x41; //A Control key code
        public const int C = 0x43; //A Control key code

        public VM()
        {
            //get a handler to the Lync Client, then subscribe to state changes.
            _client = LyncClient.GetClient();
            _client.StateChanged += _client_StateChanged;
            SubscribetoPresenceIfSignedIn(_client.State);
        }


        void _client_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
            SubscribetoPresenceIfSignedIn(e.NewState);
        }

        private void SubscribetoPresenceIfSignedIn(ClientState state)
        {
            if (state == ClientState.SignedIn)
            {
                SelfSipAddress = _client.Self.Contact.Uri;
                _client.Self.Contact.ContactInformationChanged += Contact_ContactInformationChanged;
            }
            else
            {
                //remove event handler (i.e. from previous logins etc)
                _client.Self.Contact.ContactInformationChanged -= Contact_ContactInformationChanged;
            }
        }

        private void Contact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            Debug.WriteLine("Contact information changed");
            if (e.ChangedContactInformation.Contains(ContactInformationType.Activity) || e.ChangedContactInformation.Contains(ContactInformationType.Availability))
            {
                var activity = _client.Self.Contact.GetContactInformation(ContactInformationType.ActivityId);
                ContactAvailability availability = (ContactAvailability)_client.Self.Contact.GetContactInformation(ContactInformationType.Availability);
                if (availability == ContactAvailability.Busy && activity.ToString().ToLower() == "on-the-phone")
                {
                    if (!_pauseEnabled)
                    {
                        TriggerPauseButton(activity.ToString());
                        _pauseEnabled = true;
                        UpdateStatus();

                    }
                }
                else
                {
                    if (_pauseEnabled)
                    {
                        TriggerPauseButton(activity.ToString());
                        _pauseEnabled = false;
                        UpdateStatus();
                    }
                }
            }
        }

        private void UpdateStatus()
        {
            var statusText = _pauseEnabled ? "*PAUSED*" : string.Empty;
            PauseStatus = statusText;
            OnPropertyChanged("PauseStatus");
        }

        private void TriggerPauseButton(string p)
        {
            byte playPause = 0xB3;
            PressKey(playPause);
        }

        private void PressKey(byte keyCode)
        {
            const int KEYEVENTF_EXTENDEDKEY = 0x1;
            const int KEYEVENTF_KEYUP = 0x2;
            keybd_event(keyCode, 0x45, KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(keyCode, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }

        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }
    }
}
