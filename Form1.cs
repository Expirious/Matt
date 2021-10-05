using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.IO;
using System.Media;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace Matt
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer Matt = new SpeechSynthesizer();
        SpeechRecognitionEngine startlistening = new SpeechRecognitionEngine();

        DateTime TimeNow = DateTime.Now;
        SoundPlayer sd = new SoundPlayer();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pictureBox1.ImageLocation = @"AI.gif";

            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"DefaultCommands.txt")))));
            _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Default_SpeechRecognized);
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);

            Matt.SelectVoiceByHints(VoiceGender.Male);
            Matt.SpeakAsync("Maybe it's the last time you use me");
        }

        private void Default_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text;

            if (speech == "How are you")
            {
                Matt.SpeakAsync("I am working normaly");
            }
            if (speech == "What time is it")
            {
                Matt.SpeakAsync(DateTime.Now.ToString("h mm tt"));
            }
            if (speech == "Stop talking")
            {
                Matt.SpeakAsyncCancelAll();
            }
            if (speech == "Close")
            {
                Close();
            }

            if (speech == "Show Command list")
            {
                string[] commands = (File.ReadAllLines(@"Commands.txt"));
                LstCommands.Items.Clear();
                LstCommands.SelectionMode = SelectionMode.None;
                LstCommands.Visible = true;
                int commandNumber = 1;
                foreach (string command in commands)
                {
                    if (command != "")
                    {
                        LstCommands.Items.Add(commandNumber + "." + command);
                        commandNumber++;
                    }
                    else
                    {
                        LstCommands.Items.Add("");
                    }
                }
            }
            if (speech == "Show music list")
            {
                string[] MusicList = (File.ReadAllLines(@"MusicList.txt"));
                LstCommands.Items.Clear();
                LstCommands.SelectionMode = SelectionMode.None;
                LstCommands.Visible = true;
                int musicNumber = 1;
                foreach (string Music in MusicList)
                {
                    if (Music != "")
                    {
                        LstCommands.Items.Add(musicNumber + "." + Music);
                        musicNumber++;
                    }
                    else
                    {
                        LstCommands.Items.Add("");
                    }
                }
            }
            if (speech == "Show open list")
            {
                string[] OpenList = (File.ReadAllLines(@"OpenList.txt"));
                LstCommands.Items.Clear();
                LstCommands.SelectionMode = SelectionMode.None;
                LstCommands.Visible = true;
                int openNumber = 1;
                foreach (string Open in OpenList)
                {
                    if (Open != "")
                    {
                        LstCommands.Items.Add(openNumber + "." + Open);
                        openNumber++;
                    }
                    else
                    {
                        LstCommands.Items.Add("");
                    }
                }
            }
            if (speech == "Hide list")
            {
                LstCommands.Visible = false;
            }

            if (speech == "Switch voice to male")
            {
                Matt.SelectVoiceByHints(VoiceGender.Male);
                Matt.SpeakAsync("I just switched my voice to male");
            }
            if (speech == "Switch voice to female")
            {
                Matt.SelectVoiceByHints(VoiceGender.Female);
                Matt.SpeakAsync("I just switched my voice to female");
            }

            if (speech == "Stop the music")
            {
                sd.Stop();
            }

            //Opening
            {
                if (speech == "Open Chrome")
                {
                    Process.Start("Chrome");
                }
                if (speech == "Open youtube")
                {
                    Process.Start("Chrome", "http://www.youtube.com");
                }
                if (speech == "Open lichess")
                {
                    Process.Start("Chrome", "http://lichess.org");
                }
                if (speech == "Open 10fastfingers")
                {
                    Process.Start("Chrome", "http://10ff.net");
                }
                if (speech == "Open League of Legends")
                {
                    Process.Start(@"C:\Riot Games\League of Legends\LeagueClient.exe");
                }
                if (speech == "Open Visual Studio")
                {
                    Process.Start(@"C:\Program Files\Microsoft Visual Studio\2022\Preview\Common7\IDE\devenv");
                }
                if (speech == "Open Minecraft")
                {
                    Process.Start(@"C:\Users\HP-Laptop\AppData\Roaming\.minecraft\TLauncher");
                }
                if (speech == "Open Discord")
                {
                    Process.Start(@"C:\Users\HP-Laptop\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Discord Inc\Discord");
                }
                if (speech == "Open Steam")
                {
                    Process.Start(@"C:\Program Files (x86)\Steam\steam");
                }
                if (speech == "Open Paladins")
                {
                    Process.Start(@"C:\Program Files (x86)\Steam\steam", "steam://rungameid/444090");
                }
                if (speech == "Open Spotify")
                {
                    Process.Start("Spotify");
                }
            }
            

            //Closing
            {
                if (speech == "Close Chrome")
                {
                    foreach (Process process in Process.GetProcesses())
                    {
                        if (process.MainWindowTitle.Contains("Chrome"))
                        {
                            process.Kill();
                            break;
                        }
                    }
                }
                if (speech == "Close League of Legends")
                {
                    Task.Run(() =>
                    {
                        foreach (Process process in Process.GetProcesses())
                        {
                            if (process.MainWindowTitle.Contains("League of Legends"))
                            {
                                process.CloseMainWindow();
                                Thread.Sleep(500);
                                SendKeys.Send("{ENTER}");
                                break;
                            }
                        }
                    });
                }
                if (speech == "Close Minecraft")
                {
                    Task.Run(() =>
                    {
                        for (int i = 0; i < 2; i++)
                        {
                            foreach (Process process in Process.GetProcesses())
                            {
                                if (process.MainWindowTitle.Contains("Minecraft") || process.MainWindowTitle.Contains("TLauncher"))
                                {
                                    process.CloseMainWindow();
                                    break;
                                }
                            }
                            Thread.Sleep(500);
                        }
                    });
                }
                if (speech == "Close Visual Studio")
                {
                    Task.Run(() =>
                    {
                        foreach (Process process in Process.GetProcesses())
                        {
                            if (process.MainWindowTitle.Contains("Visual Studio"))
                            {
                                process.CloseMainWindow();
                                Thread.Sleep(200);
                                SendKeys.Send("{ENTER}");
                                break;
                            }
                        }
                    });
                }
                if (speech == "Close Discord")
                {
                    foreach (Process process in Process.GetProcesses())
                    {
                        if (process.MainWindowTitle.Contains("Discord"))
                        {
                            process.Kill();
                            break;
                        }
                    }
                }
                if (speech == "Close Steam")
                {
                    foreach (Process process in Process.GetProcesses())
                    {
                        if (process.MainWindowTitle.Contains("Steam"))
                        {
                            process.CloseMainWindow();
                            break;
                        }
                    }
                }
                if (speech == "Close Paladins")
                {
                    Task.Run(() =>
                    {
                        foreach (Process process in Process.GetProcesses())
                        {
                            if (process.MainWindowTitle.Contains("Paladins"))
                            {
                                process.CloseMainWindow();
                                break;
                            }
                        }
                    });
                }
                if (speech == "Close Spotify")
                {
                    foreach (Process process in Process.GetProcesses())
                    {
                        if (process.MainWindowTitle.Contains("Spotify"))
                        {
                            Matt.SpeakAsync("Closing Spotify");
                            process.CloseMainWindow();
                            break;
                        }
                    }
                }
            }

            if (speech == "Shut Down")
            {
                Process.Start("shutdown", "/s /t 0");
            }
            if (speech == "I Go Sleep")
            {
                Matt.SpeakAsync("I'll shut down the pc in 45 minutes");

                Task.Run(() =>
                {
                    Thread.Sleep(45 * 60000);
                    Process.Start("shutdown", "/s /t 0");
                });
            }

            //if(speech == "Hide")
            //{
            //    this.Close();
            //    Thread th = new Thread(opennewform);
            //    th.SetApartmentState(ApartmentState.STA);
            //    th.Start();
            //}
        }

        private void opennewform()
        {
            //Application.Run(new Form2());
        }

        private void LstCommands_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
