/*
 * Happy Hacking Keycode
 * Copyright (c) 2020 silverintegral
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */


using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace hhkc
{
	public partial class SetForm : Form
	{
		#region Win32 Constants
		protected const int VK_SPACE = 0x20;
		protected const int VK_ESCAPE = 0x1B;
		protected const int VK_LEFT = 0x25;
		protected const int VK_UP = 0x26;
		protected const int VK_RIGHT = 0x27;
		protected const int VK_DOWN = 0x28;
		protected const int VK_POWER = 0xFF;

		protected const int VK_0 = 0x30;
		protected const int VK_1 = 0x31;
		protected const int VK_2 = 0x32;
		protected const int VK_3 = 0x33;
		protected const int VK_4 = 0x34;
		protected const int VK_5 = 0x35;
		protected const int VK_6 = 0x36;
		protected const int VK_7 = 0x37;
		protected const int VK_8 = 0x38;
		protected const int VK_9 = 0x39;

		protected const int VK_OEM_MINUS = 0xBD; // -_
		protected const int VK_OEM_EQUAL = 0xBB; // =+
		protected const int VK_OEM_1 = 0xBA; // ;:
		protected const int VK_OEM_7 = 0xDE; // '"
		protected const int VK_DECIMAL = 0x6E; // テンキー .
		protected const int VK_OEM_4 = 0xDB; // [{
		protected const int VK_OEM_6 = 0xDD; // ]}
		protected const int VK_OEM_2 = 0xBF; // /?
		protected const int VK_OEM_COMMA = 0xBC; // ,<
		protected const int VK_OEM_PERIOD = 0xBE; // .>
		protected const int VK_OEM_5 = 0xDC; // \|
		protected const int VK_OEM_3 = 0xC0; // `~

		protected const int VK_A = 0x41;
		protected const int VK_B = 0x42;
		protected const int VK_C = 0x43;
		protected const int VK_D = 0x44;
		protected const int VK_E = 0x45;
		protected const int VK_F = 0x46;
		protected const int VK_G = 0x47;
		protected const int VK_H = 0x48;
		protected const int VK_I = 0x49;
		protected const int VK_J = 0x4A;
		protected const int VK_K = 0x4B;
		protected const int VK_L = 0x4C;
		protected const int VK_M = 0x4D;
		protected const int VK_N = 0x4E;
		protected const int VK_O = 0x4F;
		protected const int VK_P = 0x50;
		protected const int VK_Q = 0x51;
		protected const int VK_R = 0x52;
		protected const int VK_S = 0x53;
		protected const int VK_T = 0x54;
		protected const int VK_U = 0x55;
		protected const int VK_V = 0x56;
		protected const int VK_W = 0x57;
		protected const int VK_X = 0x58;
		protected const int VK_Y = 0x59;
		protected const int VK_Z = 0x5A;

		protected const int VK_F1 = 0x70;
		protected const int VK_F2 = 0x71;
		protected const int VK_F3 = 0x72;
		protected const int VK_F4 = 0x73;
		protected const int VK_F5 = 0x74;
		protected const int VK_F6 = 0x75;
		protected const int VK_F7 = 0x76;
		protected const int VK_F8 = 0x77;
		protected const int VK_F9 = 0x78;
		protected const int VK_F10 = 0x79;
		protected const int VK_F11 = 0x7A;
		protected const int VK_F12 = 0x7B;

		protected const int VK_LSHIFT = 0xA0;
		protected const int VK_RSHIFT = 0xA1;
		protected const int VK_LCONTROL = 0xA2;
		protected const int VK_RCONTROL = 0xA3;
		protected const int VK_LMENU = 0xA4;
		protected const int VK_RMENU = 0xA5;
		protected const int VK_LWIN = 0x5B;
		protected const int VK_RWIN = 0x5C;

		protected const int VK_LBUTTON = 0x01;
		protected const int VK_RBUTTON = 0x02;
		protected const int VK_MBUTTON = 0x04;
		protected const int VK_XBUTTON1 = 0x05;
		protected const int VK_XBUTTON2 = 0x06;
		#endregion

		static NotifyIcon TrayIcon = null;
		static long ToastCloseTime;

		HookKey HookKey = null;
		static InputKey KeyInput = null;
		static int ModStat = 0;

		static int CtrlDown = 0;
		static int ShiftDown = 0;
		static int AltDown = 0;
		static int WinDown = 0;
		static bool EscDown = false;
		static bool EscUpTime = false;

		static int MacroStat = 0;
		static String MacroStr;

		static String[] MacroQ = new String[] { };
		static String[] MacroW = new String[] { };
		static String[] MacroE = new String[] { };
		static String[] MacroR = new String[] { };
		static String[] MacroT = new String[] { };
		static String[] MacroY = new String[] { };
		static String[] MacroU = new String[] { };
		static String[] MacroI = new String[] { };
		static String[] MacroO = new String[] { };
		static String[] MacroP = new String[] { };
		static String[] Macro1 = new String[] { };
		static String[] Macro2 = new String[] { };
		static String[] MacroExec = new String[] { };
		
		static String AppZ = "";
		static String AppX = "";
		static String AppC = "";
		static String AppV = "";
		static String AppB = "";
		static String AppN = "";
		static String AppM = "";
		static String App1 = "";
		static String App2 = "";
		static String App3 = "";


		public SetForm()
		{
			InitializeComponent();
			InitTrayIcon();
			Hide();
			LoadMacro();
			LoadApp();

			HookKey = new HookKey();
			HookKey.KeyDownEvent += KeyDownEvent;
			HookKey.KeyUpEvent += KeyUpEvent;
			HookKey.Hook();

			KeyInput = new InputKey();
		}


		public void InitTrayIcon()
		{
			TrayIcon = new NotifyIcon();
			TrayIcon.Icon = Properties.Resources.icon_tray;
			TrayIcon.Text = "Happy Hacking Keycode";
			TrayIcon.BalloonTipIcon = ToolTipIcon.Info;
			TrayIcon.BalloonTipTitle = "HHKC";
			TrayIcon.DoubleClick += new EventHandler(TrayIconClick);
			TrayIcon.Visible = true;

			ContextMenuStrip menu = new ContextMenuStrip();
			ToolStripMenuItem menuItem = new ToolStripMenuItem();
			menuItem.Text = "&終了";
			menuItem.Click += new EventHandler(TrayIconMenu_Exit);
			menu.Items.Add(menuItem);

			TrayIcon.ContextMenuStrip = menu;
		}


		private void TrayIconClick(object sender, EventArgs e)
		{
			this.WindowState = FormWindowState.Normal;
			this.ShowInTaskbar = true;
		}


		private void TrayIconMenu_Exit(object sender, EventArgs e)
		{
			TrayIcon.Visible = false;
			HookKey.UnHook();
			Application.Exit();
		}


		private void SetFormClosing(object sender, FormClosingEventArgs e)
		{
			if (e.CloseReason == CloseReason.UserClosing)
			{
				e.Cancel = true;

				this.WindowState = FormWindowState.Minimized;
				this.ShowInTaskbar = false;
			}
		}


		static private void ShowToast(String msg)
		{
			ToastCloseTime = GetUnixTime() + 3;

			TrayIcon.Visible = false;
			TrayIcon.Visible = true;

			TrayIcon.BalloonTipText = msg;
			TrayIcon.ShowBalloonTip(3000);

			Task.Run(() =>
			{
				long stdtime = ToastCloseTime;

				while (GetUnixTime() > ToastCloseTime)
				{
					Thread.Sleep(300);

					if (stdtime != ToastCloseTime)
						return;
				}

				TrayIcon.Visible = false;
				TrayIcon.Visible = true;
			});
		}


		public static long GetUnixTime()
		{
			var dto = new DateTimeOffset(DateTime.Now.Ticks, new TimeSpan(+09, 00, 00));
			return dto.ToUnixTimeSeconds();
		}


		private static void SaveMacro()
		{
			System.IO.DirectoryInfo pdir = System.IO.Directory.GetParent(Application.UserAppDataPath);
			StreamWriter sw = new StreamWriter(pdir.FullName + @"\macro.conf", false, Encoding.ASCII);

			try
			{
				sw.Write(string.Join(",", MacroQ) + "\n");
				sw.Write(string.Join(",", MacroW) + "\n");
				sw.Write(string.Join(",", MacroE) + "\n");
				sw.Write(string.Join(",", MacroR) + "\n");
				sw.Write(string.Join(",", MacroT) + "\n");
				sw.Write(string.Join(",", MacroY) + "\n");
				sw.Write(string.Join(",", MacroU) + "\n");
				sw.Write(string.Join(",", MacroI) + "\n");
				sw.Write(string.Join(",", MacroO) + "\n");
				sw.Write(string.Join(",", MacroP) + "\n");
				sw.Write(string.Join(",", Macro1) + "\n");
				sw.Write(string.Join(",", Macro2) + "\n");
			}
			finally
			{
				sw.Close();
			}
		}


		private static void LoadMacro()
		{
			System.IO.DirectoryInfo pdir = System.IO.Directory.GetParent(Application.UserAppDataPath);

			if (!File.Exists(pdir.FullName + @"\macro.conf"))
				return;

			StreamReader sr = new StreamReader(pdir.FullName + @"\macro.conf");

			try
			{
				MacroQ = sr.ReadLine().Split(',');
				MacroW = sr.ReadLine().Split(',');
				MacroE = sr.ReadLine().Split(',');
				MacroR = sr.ReadLine().Split(',');
				MacroT = sr.ReadLine().Split(',');
				MacroY = sr.ReadLine().Split(',');
				MacroU = sr.ReadLine().Split(',');
				MacroI = sr.ReadLine().Split(',');
				MacroO = sr.ReadLine().Split(',');
				MacroP = sr.ReadLine().Split(',');
				Macro1 = sr.ReadLine().Split(',');
				Macro2 = sr.ReadLine().Split(',');
			}
			finally
			{
				sr.Close();
			}
		}


		private static void SaveApp()
		{
			System.IO.DirectoryInfo pdir = System.IO.Directory.GetParent(Application.UserAppDataPath);
			StreamWriter sw = new StreamWriter(pdir.FullName + @"\app.conf", false, Encoding.ASCII);

			try
			{
				sw.Write(AppZ + "\n");
				sw.Write(AppX + "\n");
				sw.Write(AppC + "\n");
				sw.Write(AppV + "\n");
				sw.Write(AppB + "\n");
				sw.Write(AppN + "\n");
				sw.Write(AppM + "\n");
				sw.Write(App1 + "\n");
				sw.Write(App2 + "\n");
				sw.Write(App3 + "\n");
			}
			finally
			{
				sw.Close();
			}
		}


		private static void LoadApp()
		{
			System.IO.DirectoryInfo pdir = System.IO.Directory.GetParent(Application.UserAppDataPath);

			if (!File.Exists(pdir.FullName + @"\app.conf"))
				return;

			StreamReader sr = new StreamReader(pdir.FullName + @"\app.conf");

			try
			{
				AppZ = sr.ReadLine();
				AppX = sr.ReadLine();
				AppC = sr.ReadLine();
				AppV = sr.ReadLine();
				AppB = sr.ReadLine();
				AppN = sr.ReadLine();
				AppM = sr.ReadLine();
				App1 = sr.ReadLine();
				App2 = sr.ReadLine();
				App3 = sr.ReadLine();
			}
			finally
			{
				sr.Close();
			}
		}


		private static bool KeyDownEvent(object sender, HookKey.OriginalKeyEventArg e)
		{
			//string log = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " Keydown " + e.KeyCode + Environment.NewLine;
			//System.IO.File.AppendAllText("D:/log.txt", log);

			if (e.KeyCode == VK_LCONTROL)
				CtrlDown |= 1;

			if (e.KeyCode == VK_RCONTROL)
				CtrlDown |= 2;

			if (e.KeyCode == VK_LSHIFT)
				ShiftDown |= 1;

			if (e.KeyCode == VK_RSHIFT)
				ShiftDown |= 2;

			if (e.KeyCode == VK_LMENU)
				AltDown |= 1;

			if (e.KeyCode == VK_RMENU)
				AltDown |= 2;

			if (e.KeyCode == VK_LWIN)
				WinDown |= 1;

			if (e.KeyCode == VK_RWIN)
				WinDown |= 2;


			if (MacroStat > 0)
			{
				if (MacroStat == 2)
				{
					if (e.KeyCode == VK_SPACE)
					{
						MacroStat = 3;
						return false;
					}
					else
					{
						if (e.KeyCode == VK_LSHIFT || e.KeyCode == VK_RSHIFT
							|| e.KeyCode == VK_LCONTROL || e.KeyCode == VK_RCONTROL
							|| e.KeyCode == VK_LMENU || e.KeyCode == VK_RMENU
							|| e.KeyCode == VK_LWIN || e.KeyCode == VK_RWIN)
							MacroStr += (1000 + e.KeyCode).ToString() + ",";
						else
							MacroStr += e.KeyCode.ToString() + ",";
					}

					return true;
				}
				else if (MacroStat == 3)
				{
					// マクロ記録終了（場所選択）
					if (e.KeyCode == VK_Q)
					{
						MacroQ = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_W)
					{
						MacroW = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_E)
					{
						MacroE = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_R)
					{
						MacroR = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_T)
					{
						MacroT = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_Y)
					{
						MacroY = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_U)
					{
						MacroU = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_I)
					{
						MacroI = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_O)
					{
						MacroO = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_P)
					{
						MacroP = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_OEM_4) // [
					{
						Macro1 = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}
					else if (e.KeyCode == VK_OEM_6) // ]
					{
						Macro2 = MacroStr.Trim(',').Split(',');
						MacroStat = 4;
					}

					if (MacroStat == 4)
					{
						ShowToast("マクロ登録が完了しました");
						MacroStat = 0;
						ModStat = 0;
						SaveMacro();
						return false;
					}
					
					if (e.KeyCode == VK_ESCAPE)
					{
						ShowToast("マクロ登録をキャンセルしました");
						MacroStat = 0;
						ModStat = 0;
						return false;
					}
				}

				return false;
			}


			if (e.KeyCode == VK_SPACE)
			{
				if (EscDown || EscUpTime)
				{
					// ESCと同時押し、もくはく0.5秒以内で無効化
					ModStat = -1;
				}
				else if (ShiftDown > 0 || CtrlDown > 0 || AltDown > 0 || WinDown > 0)
				{
					// 事前に装飾で無効化
					ModStat = -1;
				}
				else if (ModStat == 1)
				{
					// MODが使われる事なく連打が始まったらキャンセル
					ModStat = -1;
				}
				else if (ModStat == 0)
				{
					// MODをON
					ModStat = 1;
					return false;
				}
			}

			if (e.KeyCode == VK_ESCAPE)
			{
				EscDown = true;
			}

			if (ModStat > 0)
			{
				// ファンクション簡単入力
				if (e.KeyCode >= VK_1 && e.KeyCode <= VK_9)
				{
					// F1-F9
					KeyInput.KeyDown(e.KeyCode + 63);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_0)
				{
					// F10
					KeyInput.KeyDown(VK_F10);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_OEM_MINUS)
				{
					// F11
					KeyInput.KeyDown(VK_F11);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_OEM_EQUAL)
				{
					// F12
					KeyInput.KeyDown(VK_F12);
					ModStat = 2;
					return false;
				}

				// 数字簡単入力
				if (e.KeyCode == VK_A)
				{
					// 1
					KeyInput.KeyDown(VK_1);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_S)
				{
					// 2
					KeyInput.KeyDown(VK_2);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_D)
				{
					// 3
					KeyInput.KeyDown(VK_3);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_F)
				{
					// 4
					KeyInput.KeyDown(VK_4);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_G)
				{
					// 5
					KeyInput.KeyDown(VK_5);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_H)
				{
					// 6
					KeyInput.KeyDown(VK_6);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_J)
				{
					// 7
					KeyInput.KeyDown(VK_7);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_K)
				{
					// 8
					KeyInput.KeyDown(VK_8);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_L)
				{
					// 9
					KeyInput.KeyDown(VK_9);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_OEM_1)
				{
					// 0
					KeyInput.KeyDown(VK_0);
					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_OEM_7)
				{
					// .
					KeyInput.KeyDown(VK_DECIMAL);
					ModStat = 2;
					return false;
				}

				// プログラム実行
				if (e.KeyCode == VK_Z)
				{
					if (AppZ.Length > 0)
						Process.Start(AppZ);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_X)
				{
					if (AppX.Length > 0)
						Process.Start(AppX);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_C)
				{
					if (AppC.Length > 0)
						Process.Start(AppC);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_V)
				{
					if (AppV.Length > 0)
						Process.Start(AppV);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_B)
				{
					if (AppB.Length > 0)
						Process.Start(AppB);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_N)
				{
					if (AppN.Length > 0)
						Process.Start(AppN);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_M)
				{
					if (AppM.Length > 0)
						Process.Start(AppM);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_OEM_COMMA)
				{
					if (App1.Length > 0)
						Process.Start(App1);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_OEM_PERIOD)
				{
					if (App2.Length > 0)
						Process.Start(App2);

					ModStat = 2;
					return false;
				}
				else if (e.KeyCode == VK_OEM_2)
				{
					if (App3.Length > 0)
						Process.Start(App3);

					ModStat = 2;
					return false;
				}

				// マクロ記録
				if (e.KeyCode == VK_ESCAPE)
				{
					MacroStat = 1;
					MacroStr = "";
					return false;
				}

				// マクロ再生
				if (e.KeyCode == VK_Q)
				{
					MacroExec = MacroQ;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_W)
				{
					MacroExec = MacroW;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_E)
				{
					MacroExec = MacroE;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_R)
				{
					MacroExec = MacroR;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_T)
				{
					MacroExec = MacroT;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_Y)
				{
					MacroExec = MacroY;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_U)
				{
					MacroExec = MacroU;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_I)
				{
					MacroExec = MacroI;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_O)
				{
					MacroExec = MacroO;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_P)
				{
					MacroExec = MacroP;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_OEM_4)
				{
					MacroExec = Macro1;
					MacroStat = 9;
				}
				else if (e.KeyCode == VK_OEM_6)
				{
					MacroExec = Macro2;
					MacroStat = 9;
				}

				if (MacroStat == 9)
				{
					foreach (String strkey in MacroExec)
					{
						if (strkey.Length == 0)
							continue;

						int key = Int32.Parse(strkey);
						if (key > 2000)
						{
							key -= 2000;
							KeyInput.KeyUp(key);
						}
						else if (key > 1000)
						{
							key -= 1000;
							KeyInput.KeyDown(key);
						}
						else
						{
							KeyInput.KeyDown(key);
							KeyInput.KeyUp(key);
						}
					}

					MacroStat = 0;
					ModStat = 2;
					return false;
				}

				ModStat = 2;
			}

			// POWERボタンの有効利用
			// (Macintoshモードのみ)
			if (e.KeyCode == VK_POWER)
			{
				if (ShiftDown > 0)
					SystemCtrl.Suspend();
				else
					SystemCtrl.MonitorOff();

				return false;
			}

			return true;
		}


		private static bool KeyUpEvent(object sender, HookKey.OriginalKeyEventArg e)
		{
			if (e.KeyCode == VK_LCONTROL)
				CtrlDown ^= 1;

			if (e.KeyCode == VK_RCONTROL)
				CtrlDown ^= 2;

			if (e.KeyCode == VK_LSHIFT)
				ShiftDown ^= 1;

			if (e.KeyCode == VK_RSHIFT)
				ShiftDown ^= 2;

			if (e.KeyCode == VK_LMENU)
				AltDown ^= 1;

			if (e.KeyCode == VK_RMENU)
				AltDown ^= 2;

			if (e.KeyCode == VK_LWIN)
				WinDown ^= 1;

			if (e.KeyCode == VK_RWIN)
				WinDown ^= 2;


			if (e.KeyCode == VK_ESCAPE)
			{
				EscDown = false;

				if (ModStat == 0)
				{
					var task = Task.Run(() =>
					{
						EscUpTime = true;
						Thread.Sleep(600);
						EscUpTime = false;
					});
				}
			}

			if (MacroStat > 0)
			{
				if (MacroStat == 1)
				{
					if (e.KeyCode == VK_Z)
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							AppZ = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								AppZ = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_X)
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							AppX = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								AppX = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_C)
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							AppC = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								AppC = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_V)
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							AppV = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								AppV = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_B)
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							AppB = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								AppB = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_N)
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							AppN = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								AppN = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_M)
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							AppM = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								AppM = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_OEM_COMMA) // ,
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							App1 = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								App1 = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_OEM_PERIOD) // .
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							App2 = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								App2 = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_OEM_2) // /
					{
						MacroStat = 0;
						ModStat = 1;

						if (ShiftDown > 0)
						{
							ShowToast("アプリを削除しました");
							App3 = "";
							SaveApp();
						}
						else
						{
							ShowToast("アプリ登録");

							OpenFileDialog Dlg = new OpenFileDialog
							{
								InitialDirectory = @"C:",
								Title = "アプリの登録"
							};

							if (Dlg.ShowDialog() == DialogResult.OK)
							{
								App3 = Dlg.FileName;
								SaveApp();
								ShowToast("アプリ登録が完了しました");
							}
							else
							{
								ShowToast("アプリ登録をキャンセルしました");
							}

							Dlg.Dispose();
						}

						return false;
					}
					else if (e.KeyCode == VK_SPACE)
					{
						// スペースが上がってからマクロの記録開始
						MacroStat = 2;
						ShowToast("マクロ登録の開始");
						return false;
					}
				}
				else if (MacroStat == 2)
				{
					// 装飾キーはUpも記録
					if (e.KeyCode == VK_LSHIFT || e.KeyCode == VK_RSHIFT
						|| e.KeyCode == VK_LCONTROL || e.KeyCode == VK_RCONTROL
						|| e.KeyCode == VK_LMENU || e.KeyCode == VK_RMENU)
					{
						MacroStr += (2000 + e.KeyCode).ToString() + ",";
						return true;
					}
				}
				else if (MacroStat == 3)
				{
					if (e.KeyCode == VK_SPACE)
					{
						// マクロの記録はスペースも通常通り行う
						KeyInput.KeyDown(VK_SPACE);
						KeyInput.KeyUp(VK_SPACE);
						MacroStr += e.KeyCode.ToString() + ",";
						MacroStat = 2;
						return false;
					}
				}

				return false;
			}


			if (e.KeyCode == VK_SPACE)
			{
				if (ModStat == 1)
				{
					KeyInput.KeyDown(VK_SPACE);
					KeyInput.KeyUp(VK_SPACE);
				}

				ModStat = 0;
				return false;
			}

			if (ModStat > 0)
			{
				// ファンクション簡単入力
				if (e.KeyCode >= VK_1 && e.KeyCode <= VK_9)
				{
					// F1-F9
					KeyInput.KeyUp(e.KeyCode + 63);
					return false;
				}
				else if (e.KeyCode == VK_0)
				{
					// F10
					KeyInput.KeyUp(VK_F10);
					return false;
				}
				else if (e.KeyCode == VK_OEM_MINUS)
				{
					// F11
					KeyInput.KeyUp(VK_F11);
					return false;
				}
				else if (e.KeyCode == VK_OEM_EQUAL)
				{
					// F12
					KeyInput.KeyUp(VK_F12);
					return false;
				}

				// 数字簡単入力
				if (e.KeyCode == VK_A)
				{
					// 1
					KeyInput.KeyUp(VK_1);
					return false;
				}
				else if (e.KeyCode == VK_S)
				{
					// 2
					KeyInput.KeyUp(VK_2);
					return false;
				}
				else if (e.KeyCode == VK_D)
				{
					// 3
					KeyInput.KeyUp(VK_3);
					return false;
				}
				else if (e.KeyCode == VK_F)
				{
					// 4
					KeyInput.KeyUp(VK_4);
					return false;
				}
				else if (e.KeyCode == VK_G)
				{
					// 5
					KeyInput.KeyUp(VK_5);
					return false;
				}
				else if (e.KeyCode == VK_H)
				{
					// 6
					KeyInput.KeyUp(VK_6);
					return false;
				}
				else if (e.KeyCode == VK_J)
				{
					// 7
					KeyInput.KeyUp(VK_7);
					return false;
				}
				else if (e.KeyCode == VK_K)
				{
					// 8
					KeyInput.KeyUp(VK_8);
					return false;
				}
				else if (e.KeyCode == VK_L)
				{
					// 9
					KeyInput.KeyUp(VK_9);
					return false;
				}
				else if (e.KeyCode == VK_OEM_1)
				{
					// 0
					KeyInput.KeyUp(VK_0);
					return false;
				}
				else if (e.KeyCode == VK_OEM_7)
				{
					// .
					KeyInput.KeyUp(VK_DECIMAL);
					return false;
				}

				/*
				if (e.KeyCode == VK_OEM_5) // \
				{
					return false;
				}

				if (e.KeyCode == VK_OEM_3) // `
				{
					return false;
				}
				*/
			}

			return true;
		}
	}


	public class HookKey
	{
		#region Win32 Constants
		protected const int WH_KEYBOARD_LL = 0x000D;
		protected const int WM_KEYDOWN = 0x0100;
		protected const int WM_KEYUP = 0x0101;
		protected const int WM_SYSKEYDOWN = 0x0104;
		protected const int WM_SYSKEYUP = 0x0105;
		#endregion

		#region Win32API Structures
		[StructLayout(LayoutKind.Sequential)]
		public class KBDLLHOOKSTRUCT
		{
			public uint vkCode;
			public uint scanCode;
			public KBDLLHOOKSTRUCTFlags flags;
			public uint time;
			public UIntPtr dwExtraInfo;
		}

		[Flags]
		public enum KBDLLHOOKSTRUCTFlags : uint
		{
			KEYEVENTF_EXTENDEDKEY = 0x0001,
			KEYEVENTF_KEYUP = 0x0002,
			KEYEVENTF_SCANCODE = 0x0008,
			KEYEVENTF_UNICODE = 0x0004,
		}
		#endregion

		#region Win32 Methods
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr SetWindowsHookEx(int idHook, KeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool UnhookWindowsHookEx(IntPtr hhk);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr GetModuleHandle(string lpModuleName);
		#endregion

		#region Delegate
		private delegate IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
		#endregion

		#region Fields
		private IntPtr hookId = IntPtr.Zero;
		#endregion

		#region InputEvent
		public class OriginalKeyEventArg : EventArgs
		{
			public int KeyCode { get; }

			public OriginalKeyEventArg(int keyCode)
			{
				KeyCode = keyCode;
			}
		}
		public delegate bool KeyEventHandler(object sender, OriginalKeyEventArg e);
		public event KeyEventHandler KeyDownEvent;
		public event KeyEventHandler KeyUpEvent;

		protected bool OnKeyDownEvent(int keyCode)
		{
			return (bool)KeyDownEvent?.Invoke(this, new OriginalKeyEventArg(keyCode));
		}

		protected bool OnKeyUpEvent(int keyCode)
		{
			return (bool)KeyUpEvent?.Invoke(this, new OriginalKeyEventArg(keyCode));
		}
		#endregion


		public void Hook()
		{
			using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
			{
				using (ProcessModule curModule = curProcess.MainModule)
				{
					hookId = SetWindowsHookEx(WH_KEYBOARD_LL, HookProcedure, GetModuleHandle(curModule.ModuleName), 0);
				}
			}
		}


		public void UnHook()
		{
			UnhookWindowsHookEx(hookId);
		}


		public IntPtr HookProcedure(int nCode, IntPtr wParam, IntPtr lParam)
		{
			if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
			{
				var kb = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
				if ((int)(kb.dwExtraInfo) != 8088)
				{
					var vkCode = (int)kb.vkCode;
					if (!OnKeyDownEvent(vkCode))
						return new IntPtr(1);
				}
			}
			else if (nCode >= 0 && (wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP))
			{
				var kb = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
				if ((int)(kb.dwExtraInfo) != 8088)
				{
					var vkCode = (int)kb.vkCode;
					if (!OnKeyUpEvent(vkCode))
						return new IntPtr(1);
				}
			}

			return CallNextHookEx(hookId, nCode, wParam, lParam);
		}
	}


	public class InputKey
	{
		#region Win32API Structures
		[StructLayout(LayoutKind.Sequential)]
		public struct MOUSEINPUT
		{
			public int dx;
			public int dy;
			public int mouseData;
			public int dwFlags;
			public int time;
			public int dwExtraInfo;
		};

		[StructLayout(LayoutKind.Sequential)]
		public struct KEYBDINPUT
		{
			public short wVk;
			public short wScan;
			public int dwFlags;
			public int time;
			public int dwExtraInfo;
		};

		[StructLayout(LayoutKind.Sequential)]
		public struct HARDWAREINPUT
		{
			public int uMsg;
			public short wParamL;
			public short wParamH;
		};

		[StructLayout(LayoutKind.Explicit)]
		public struct INPUT
		{
			[FieldOffset(0)]
			public int type;
			[FieldOffset(4)]
			public MOUSEINPUT no;
			[FieldOffset(4)]
			public KEYBDINPUT ki;
			[FieldOffset(4)]
			public HARDWAREINPUT hi;
		};
		#endregion

		#region Win32API Methods
		[DllImport("user32.dll")]
		private static extern void SendInput(int nInputs, ref INPUT pInputs, int cbsize);
		[DllImport("user32.dll", EntryPoint = "MapVirtualKeyA")]
		private static extern int MapVirtualKey(int wCode, int wMapType);
		#endregion

		#region Win32API Constants
		private const int INPUT_KEYBOARD = 1;
		private const int KEYEVENTF_KEYDOWN = 0x0;
		private const int KEYEVENTF_KEYUP = 0x2;
		private const int KEYEVENTF_EXTENDEDKEY = 0x1;
		#endregion


		public void KeyDown(int key, bool isExtend = false)
		{
			INPUT input = new INPUT
			{
				type = INPUT_KEYBOARD,
				ki = new KEYBDINPUT()
				{
					wVk = (short)key,
					wScan = (short)MapVirtualKey((short)key, 0),
					dwFlags = ((isExtend) ? (KEYEVENTF_EXTENDEDKEY) : 0x0) | KEYEVENTF_KEYDOWN,
					time = 0,
					dwExtraInfo = 8088
				},
			};

			SendInput(1, ref input, Marshal.SizeOf(input));
		}

		public void KeyUp(int key, bool isExtend = false)
		{
			INPUT input = new INPUT
			{
				type = INPUT_KEYBOARD,
				ki = new KEYBDINPUT()
				{
					wVk = (short)key,
					wScan = (short)MapVirtualKey((short)key, 0),
					dwFlags = ((isExtend) ? (KEYEVENTF_EXTENDEDKEY) : 0x0) | KEYEVENTF_KEYUP,
					time = 0,
					dwExtraInfo = 8088
				},
			};

			SendInput(1, ref input, Marshal.SizeOf(input));
		}
	}


	public class SystemCtrl
	{
		#region Win32API Constants
		private const int SC_MONITORPOWER = 0xf170;
		private const int WM_SYSCOMMAND = 0x112;
		#endregion

		[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
		static extern IntPtr SendMessage(int hWnd, uint Msg, int wParam, int lParam);


		public static void MonitorOff()
		{
			//モニター停止
			SendMessage(-1, WM_SYSCOMMAND, SC_MONITORPOWER, 2);
		}

		public static void MonitorOn()
		{
			//モニター復帰
			SendMessage(-1, WM_SYSCOMMAND, SC_MONITORPOWER, -1);
		}

		public static void Suspend()
		{
			Application.SetSuspendState(PowerState.Suspend, false, false);
		}
	}
}
