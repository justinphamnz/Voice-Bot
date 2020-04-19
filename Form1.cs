using System;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.IO.Ports;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Voice_Bot.Constants;

/**
 * Hi, this is C#.
 * 
 * If you don't understand any method or class, you can 
 * move your mouse to the key word (eg. "SpeechSynthesizer") 
 * to see its function or Ctrl + click to navigate to its place.
 * 
 * You also need to add System.Speak reference for using speech:
 * 1/ In Solution Explorer window, under Voice Bot Solution, there are several files such as Properties, 
 *    References, App.config, Form1.cs, Program.cs, etc.
 * 2/ Right click on References -> Add Reference...
 * 3/ Search (Ctrl+E) Speech
 * 4/ Click System.Speech -> OK
 * 
 */
namespace Voice_Bot  //Naming convetion: we should use the name with no space for Solution Name
{
    public partial class Form1 : Form
    {
        #region Private members
        string name = "Taylor";     //set my name
        static bool wake = false;   //set state to "Deaf"
        bool isLiked = true;        //set if I like the voice bot
        bool search = false;        //set if I want to search something in Google
        bool closeApp = false;      //set if I want to close an application 
        bool isRunning = false;     //set if the app I want to close is running
        bool exit = false;          //set if I want to exit voice bot

        //Declare Speech
        SpeechSynthesizer speech = new SpeechSynthesizer();
        Choices grammarList = new Choices();
        SpeechRecognitionEngine record = new SpeechRecognitionEngine();

        //Declare WeatherData
        static string city = "Auckland";
        WeatherData weather = new WeatherData(city);

        //Define cursor position
        static int positionY;
        static int positionX;
        int moveArea = 50; //the amount of unit your mouse will move

        //Declare Arduino port
        static SerialPort port = new SerialPort("COM5", 9600, Parity.None, 8, StopBits.One);
        static bool light = false; //set light status

        //Assign command and response
        string[] comList = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Command.txt");
        string[] resList = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Response.txt");
        string[] googleSearch = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Google search.txt");

        /**
         * Note: Command.txt and Response.txt are linking to each other by line number 
         *   (eg. line 9 in Command.txt is "how are you" and line 9 in Response.txt is "Good, and you").
         */

        public Form1()
        {
            //Define the contraints for speech recognition
            grammarList.Add(comList);
            grammarList.Add(googleSearch);
            Grammar grammar = new Grammar(new GrammarBuilder(grammarList));

            try
            {
                //Working with speech recogniser
                record.RequestRecognizerUpdate();
                record.LoadGrammar(grammar);
                record.SpeechRecognized += RecordSpeechRecognized;
                record.SetInputToDefaultAudioDevice();
                record.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch
            {
                return;
            }

            //Set voice gender to female
            speech.SelectVoiceByHints(VoiceGender.Female);

            //This sentence will be spoken at the beginning.
            speech.Speak("Hohohoho, I am Santa Claus.");

            InitializeComponent();
        }

        #endregion

        #region Methods
        public void Restart()
        {
            //open another "Voice Bot.exe"
            Process.Start(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\bin\Debug\Voice Bot.exe");

            //Exit current voice bot application
            Environment.Exit(0);
        }

        public void Say(string response)
        {
            //Speak response
            speech.Speak(response); //Text appears after speech

            //Update response in "RESPONSE" textbox.
            textBox2.AppendText(response + "\n");

            //Set state to "Deaf"
            ChangeState(false, stateLabel);
        }

        public static void ChangeState(bool nextState, Label stateLabel)
        {
            wake = nextState;

            //Update "State" label
            if (nextState)
            {
                stateLabel.Text = "State: Listening";
            }
            else
            {
                stateLabel.Text = "State: Deaf";
            }
        }

        public static bool KillProgram(string processName)
        {
            //if this app is not running
            if (Process.GetProcessesByName(processName).Length == 0)
            {
                return false;
            }
            else
            {
                foreach (Process process in Process.GetProcessesByName(processName))
                {
                    process.Kill(); //Kill process
                }

                return true;
            }
        }

        public static void UpdateMousePosition()
        {
            positionY = Cursor.Position.Y;
            positionX = Cursor.Position.X;
        }

        public static void ChangeLightStatus(string id)
        {
            //Accessing port and updating light status
            port.Open();
            port.WriteLine(id);
            port.Close();

            if (id == "A")
            {
                light = true;
            }
            else
            {
                light = false;
            }
        }

        #endregion

        #region Private Method

        /**
 * Interacting method (getting command and responding)
 * When you say something assigned in grammarList, this method will be invoked.
 * Say "Wake" to make it work ("Listening")
 * Say "Sleep" to make it stop listening ("Deaf")
 * Whenever the voice bot talks something to you, its state will change
 * from "Listening" to "Deaf" -> Stop listening.
 */
        private void RecordSpeechRecognized(object sender, SpeechRecognizedEventArgs e) //Naming convention 
        {
            //Saving your speech
            string record = e.Result.Text;

            //Find record's position in comList array to get resList[count]
            int count = Array.IndexOf(comList, record);

            //if search = true and you say something in grammarList 
            if (search)
            {
                Process.Start("https://www.google.com/search?q=" + record); //Google search
                ChangeState(false, stateLabel); //Change state to "Deaf"
                search = false;
            }

            //If closeApp = true and you say an app name in the grammarList (eg. Teams)
            if (closeApp)
            {
                if (record == "Google")
                {
                    isRunning = KillProgram("chrome");
                    //isRunning is used for announcing if this app is not running or has been closed.
                }
                else
                {
                    isRunning = KillProgram(record);
                }

                closeApp = false;
            }

            //Say "wake" to invoke the voice bot
            if (record == "wake")
            {
                ChangeState(true, stateLabel);
            }

            /**
             * DOING TASKS when wake = true and search = false.
             * From here, every response will be processed according to its type
             * (Multi responses, Date and Time, No response, One response, Exit)
             * and its function (eg. Weather), identified through Response.txt.
             */
            if (wake == true && search == false)
            {
                //Multi responses - starts with '+'
                if (resList[count].StartsWith("+"))
                {
                    HandleMultipleResponse(count);
                }
                //Date and Time - starts with '-'
                else if (resList[count].StartsWith("-"))
                {
                    HandleDateTimeResponse(count);
                }
                //No response - starts with '~' (the state will not automatically change to "Deaf")
                else if (resList[count].StartsWith("~"))
                {
                    HandleNoResponse(record);
                }
                //One response - the rest responses not including "exit", "yes", and "no"
                else if (record != "exit" && record != "yes" && record != "no")
                {
                    HandleOtherResponse(count);
                }
                //Responses including "exit", "yes", and "no"
                else
                {
                    if (record == "exit")
                    {
                        HandleExit(count);
                    }

                    if (record == "yes" && exit == true) //turn off the light and exit
                    {
                        HandleYes(count);

                    }

                    if (record == "no" && exit == true) //exit without turning off the light
                    {
                        HandleNo(count);
                    }
                }
            }

            //Update command in "TALK" textbox
            textBox1.AppendText(record + "\n");
        }

        private void HandleMultipleResponse(int responsePosition)
        {
            //Seperate multi responses by '/'
            List<string> multiRes = resList[responsePosition].
                Substring(resList[responsePosition].LastIndexOf(']') + 1).Split('/').ToList();

            string response = resList[responsePosition];

            string responseResult = resList.FirstOrDefault<string>(s => response.Contains(s));

            switch (responseResult)
            {
                case ResponseConstant.Greetings:

                    Random random = new Random();
                    Say(multiRes[random.Next(multiRes.Count())]);

                    break;
                case ResponseConstant.Like:

                    if (isLiked)
                    {
                        Say(multiRes[0]);
                    }
                    else
                    {
                        Say(multiRes[1] + name + multiRes[2]);
                    }

                    break;

                case ResponseConstant.Delete:

                    Say(multiRes[0]);
                    speech.SelectVoiceByHints(VoiceGender.Male);

                    Say(multiRes[1]);
                    speech.SelectVoiceByHints(VoiceGender.Female);

                    Say(multiRes[2]);

                    break;
                case ResponseConstant.Weather:

                    Say(multiRes[0]);
                    weather.CheckWeather();
                    Say(multiRes[1] + weather.Condition + multiRes[2] + name);

                    break;

                case ResponseConstant.Temp:

                    Say(multiRes[0]);
                    weather.CheckWeather();
                    Say(multiRes[1] + weather.Temperature.ToString() + multiRes[2]);

                    break;

                case ResponseConstant.Close:

                    if (isRunning)
                    {
                        Say(multiRes[0]);
                    }
                    else
                    {
                        Say(multiRes[1]);
                    }

                    break;

                default:
                    break;
            }
        }

        private void HandleDateTimeResponse(int responsePosition)
        {
            string response = resList[responsePosition];

            string responseResult = resList.FirstOrDefault<string>(s => response.Contains(s));

            switch (responseResult)
            {
                case ResponseConstant.Date:

                    Say(DateTime.Now.ToString("M/dd/yyyy"));

                    break;

                case ResponseConstant.Time:

                    Say(DateTime.Now.ToString("h:mm tt"));

                    break;
                default:
                    break;
            }
        }

        private void HandleNoResponse(string record)
        {
            //Should apply Constants for these conditions

            if (record == "restart" || record == "update")
            {
                Restart();
            }

            if (record == "close")
            {
                closeApp = true;
            }

            //Google search
            if (record == "search for")
            {
                search = true;
            }

            //Change cursor location
            if (record == "down")
            {
                UpdateMousePosition();
                Voice_Bot.Peripherals.Mouse.MoveToPoint(positionX, positionY += moveArea);
            }

            if (record == "up")
            {
                UpdateMousePosition();
                Voice_Bot.Peripherals.Mouse.MoveToPoint(positionX, positionY -= moveArea);
            }

            if (record == "left")
            {
                UpdateMousePosition();
                Voice_Bot.Peripherals.Mouse.MoveToPoint(positionX -= moveArea, positionY);
            }

            if (record == "right")
            {
                UpdateMousePosition();
                Voice_Bot.Peripherals.Mouse.MoveToPoint(positionX += moveArea, positionY);
            }

            //Mouse click
            if (record == "click")
            {
                Voice_Bot.Peripherals.Mouse.DoMouseClick();
            }

            //Mouse right click
            if (record == "right click")
            {
                Voice_Bot.Peripherals.Mouse.DoMouseRightClick();
            }

            //Send Keys
            if (record == "play" || record == "pause")
            {
                SendKeys.Send(" ");
            }

            if (record == "next")
            {
                SendKeys.Send("^{RIGHT}");
            }

            if (record == "last")
            {
                SendKeys.Send("^{LEFT}");
            }

            if (record == "enter")
            {
                SendKeys.Send("{ENTER}");
            }

            if (record == "change tab")
            {
                SendKeys.Send("^{TAB}");
            }

            if (record == "new tab")
            {
                SendKeys.Send("^T");
            }

            if (record == "close tab")
            {
                SendKeys.Send("^W");
            }

            if (record == "close previous tab")
            {
                SendKeys.Send("^+{TAB}");
                SendKeys.Send("^W");
            }

            if (record == "close following tab")
            {
                SendKeys.Send("^{TAB}");
                SendKeys.Send("^W");
            }

            if (record == "open secret tab")
            {
                SendKeys.Send("^+N");
            }

            /**
             * Open Arduino file (One_led) to see what will happen when 
             * port.WriteLine("A") or port.WriteLine("B") is called.
             * 
             * Note: If you do not connect sensor to your computer, 
             * it will cause bug.
             */
            if (record == "light on")
            {
                ChangeLightStatus("A");
            }

            if (record == "light off")
            {
                ChangeLightStatus("B");
            }
        }

        private void HandleOtherResponse(int responsePosition)
        {
            string response = resList[responsePosition];

            string responseResult = resList.FirstOrDefault<string>(s => response.Contains(s));

            //Apply the same approach of HandleMultipleResponse.
            //Step 1: Apply Constants 
            //Step 2: Change to Switch case

            //if (resList[count].Contains("Sleep"))
            //{
            //    ChangeState(false, stateLabel);
            //}

            //if (resList[count].Contains("Like"))
            //{
            //    isLiked = true;
            //}

            //if (resList[count].Contains("Hate"))
            //{
            //    isLiked = false;
            //}

            //if (resList[count].Contains("Nezuko"))
            //{
            //    //Open wav file
            //    System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Media\nezuko chan.wav");
            //    player.Play();
            //}

            ////Change window status
            //if (resList[count].Contains("Minimise"))
            //{
            //    this.WindowState = FormWindowState.Normal;
            //}

            //if (resList[count].Contains("Maximise"))
            //{
            //    this.WindowState = FormWindowState.Maximized;
            //}

            ////Access applications
            //if (resList[count].Contains("Google"))
            //{
            //    Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
            //}

            //if (resList[count].Contains("Spotify"))
            //{
            //    Process.Start(@"C:\Users\pc\AppData\Roaming\Spotify\Spotify.exe");
            //}

            //if (resList[count].Contains("Teams"))
            //{
            //    Process.Start(@"C:\Users\pc\AppData\Local\Microsoft\Teams\current\Teams.exe");
            //}

            ////Send the link to search bar -> then enter
            //if (resList[count].Contains("YouTube"))
            //{
            //    SendKeys.Send("https://www.youtube.com");
            //    SendKeys.Send(((char)13).ToString()); //enter
            //}

            //if (resList[count].Contains("Max screen"))
            //{
            //    SendKeys.Send("f");
            //}

            ////If response contains "["
            //if (resList[count].Contains("["))
            //{
            //    resList[count] = resList[count].Substring(resList[count].LastIndexOf(']') + 1);
            //}

            //Say(resList[count]);
        }

        private void HandleExit(int responsePosition)
        {
            string response = resList[responsePosition];

            //If the light is on, ask whether user wants to turn it off
            if (light)
            {
                Say(response.Substring(response.LastIndexOf(']') + 1));
                exit = true;
            }
            else
            {
                Say(resList[responsePosition + 1].Substring(resList[responsePosition + 1].LastIndexOf(']') + 1));
                Environment.Exit(0);
            }
        }

        private void HandleYes(int responsePosition)
        {
            string response = resList[responsePosition];

            Say(response.Substring(response.LastIndexOf(']') + 1));
            ChangeLightStatus("B");
            Environment.Exit(0);
        }

        private void HandleNo(int responsePosition)
        {
            string response = resList[responsePosition];

            Say(response.Substring(response.LastIndexOf(']') + 1));
            Environment.Exit(0);
        }

        #endregion
    }
}
