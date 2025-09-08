using System;
using System.Globalization;
using System.IO;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;      
using System.Windows.Input;
using System.Windows.Interop;       
using System.Windows.Media;         
using System.Windows.Media.Imaging;
using System.Windows.Resources;
using System.Windows.Threading;

namespace WpfAnalogClock
{
    public partial class MainWindow : Window
    {
        private readonly DispatcherTimer _tickTimer = new DispatcherTimer();
        private readonly DispatcherTimer _alarmTimer = new DispatcherTimer();

        private SoundPlayer? _player;
        private MemoryStream? _wavStream;
        private bool _isRinging = false;
        private TimeSpan? _alarmTime = null;

        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        private const int WM_NCLBUTTONDOWN = 0x00A1;
        private const int HTLEFT = 10, HTRIGHT = 11, HTTOP = 12, HTTOPLEFT = 13, HTTOPRIGHT = 14, HTBOTTOM = 15, HTBOTTOMLEFT = 16, HTBOTTOMRIGHT = 17;

        private double _baseWidth;
        private double _baseHeight;
        private double _baseMinWidth;
        private double _baseMinHeight;
        private double _scale = 1.0; // 現在の倍率

        public MainWindow()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "InitializeComponent() で例外", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }

            _baseWidth = this.Width;
            _baseHeight = this.Height;
            _baseMinWidth = this.MinWidth;
            _baseMinHeight = this.MinHeight;

            try
            {
                LoadResources();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "LoadResources() で例外", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // タイマー開始（1秒ごとに更新）
            _tickTimer.Interval = TimeSpan.FromSeconds(1);
            _tickTimer.Tick += (_, __) => UpdateClock();
            _tickTimer.Start();

            _alarmTimer.Interval = TimeSpan.FromSeconds(1);
            _alarmTimer.Tick += (_, __) => CheckAlarm();
            _alarmTimer.Start();

            UpdateClock();

            ApplyScale(1.0);
        }

        private void CtxMenu_Opened(object sender, RoutedEventArgs e)
        {
            if (MiToggleAlarm != null && AlarmPanel != null)
            {
                MiToggleAlarm.Header = (AlarmPanel.Visibility == Visibility.Visible)
                    ? "アラームを隠す"
                    : "アラームを表示";
            }

            if (MiTopmost != null)
            {
                MiTopmost.IsChecked = this.Topmost;
            }

            SetScaleMenuChecks(_scale);
        }

        private void Menu_ToggleAlarm_Click(object sender, RoutedEventArgs e)
        {
            if (AlarmPanel == null) return;

            if (AlarmPanel.Visibility == Visibility.Visible)
            {
                AlarmPanel.Visibility = Visibility.Collapsed;
            }
            else
            {
                AlarmPanel.Visibility = Visibility.Visible;
                ApplyAlarmFromText();
            }

            if (MiToggleAlarm != null)
            {
                MiToggleAlarm.Header = (AlarmPanel.Visibility == Visibility.Visible)
                    ? "アラームを隠す"
                    : "アラームを表示";
            }
        }

        private void Menu_Topmost_Click(object sender, RoutedEventArgs e)
        {
            this.Topmost = !this.Topmost;
            if (MiTopmost != null)
                MiTopmost.IsChecked = this.Topmost; 
        }

        private void Menu_Scale_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi && mi.Tag is string s && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var scale))
            {
                ApplyScale(scale);
                SetScaleMenuChecks(scale);
            }
        }

        private void ApplyScale(double scale)
        {
            _scale = scale;

            RootGrid.LayoutTransform = new ScaleTransform(scale, scale);

            this.Width = _baseWidth * scale;
            this.Height = _baseHeight * scale;

            this.MinWidth = _baseMinWidth * scale;
            this.MinHeight = _baseMinHeight * scale;
        }

        private void SetScaleMenuChecks(double scale)
        {
            // スケール候補
            (MenuItem? mi, double v)[] items =
            {
                (MiScale050, 0.5),
                (MiScale075, 0.75),
                (MiScale100, 1.0),
                (MiScale125, 1.25),
                (MiScale150, 1.5),
            };

            // 微小な差は許容
            foreach (var (mi, v) in items)
            {
                if (mi != null)
                    mi.IsChecked = Math.Abs(scale - v) < 0.0001;
            }
        }

        private void BeginResize(int hitTest)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd == IntPtr.Zero) return;
            if (Mouse.LeftButton != MouseButtonState.Pressed) return;
            SendMessage(hwnd, WM_NCLBUTTONDOWN, (IntPtr)hitTest, IntPtr.Zero);
        }
        private void Resize_Left_MouseDown(object sender, MouseButtonEventArgs e) => BeginResize(HTLEFT);
        private void Resize_Right_MouseDown(object sender, MouseButtonEventArgs e) => BeginResize(HTRIGHT);
        private void Resize_Top_MouseDown(object sender, MouseButtonEventArgs e) => BeginResize(HTTOP);
        private void Resize_Bottom_MouseDown(object sender, MouseButtonEventArgs e) => BeginResize(HTBOTTOM);
        private void Resize_TopLeft_MouseDown(object sender, MouseButtonEventArgs e) => BeginResize(HTTOPLEFT);
        private void Resize_TopRight_MouseDown(object sender, MouseButtonEventArgs e) => BeginResize(HTTOPRIGHT);
        private void Resize_BottomLeft_MouseDown(object sender, MouseButtonEventArgs e) => BeginResize(HTBOTTOMLEFT);
        private void Resize_BottomRight_MouseDown(object sender, MouseButtonEventArgs e) => BeginResize(HTBOTTOMRIGHT);

        private void LoadResources()
        {
            const string FACE = "Clock_Face-001.png";
            const string HOUR = "Clock-Hand-001h.png";
            const string MIN = "Clock-Hand-001m.png";
            const string WAV = "Clock-Alarm04-01.wav";

            ImgFace.Source = LoadBitmapWithFallback("resources", FACE, out _);
            ImgHour.Source = LoadBitmapWithFallback("resources", HOUR, out _);
            ImgMinute.Source = LoadBitmapWithFallback("resources", MIN, out _);

            _wavStream = LoadStreamWithFallback("resources", WAV, out _);
            _player = (_wavStream != null) ? new SoundPlayer(_wavStream) : null;
            _player?.Load();
        }

        private BitmapImage? LoadBitmapWithFallback(string folder, string file, out string info)
        {
            var logs = new System.Text.StringBuilder();

            if (TryLoadBitmapFromPack($@"pack://application:,,,/;component/{folder}/{file}", out var bmp1, out var e1))
            { info = $"OK component:{folder}/{file}"; return bmp1; }
            logs.AppendLine($"x component: {e1?.Message}");

            if (TryLoadBitmapFromPack($@"pack://application:,,,/{folder}/{file}", out var bmp2, out var e2))
            { info = $"OK pack:{folder}/{file}"; return bmp2; }
            logs.AppendLine($"x pack: {e2?.Message}");

            if (TryLoadBitmapFromPack($@"pack://siteoforigin:,,,/{folder}/{file}", out var bmp3, out var e3))
            { info = $"OK siteoforigin:{folder}/{file}"; return bmp3; }
            logs.AppendLine($"x siteoforigin: {e3?.Message}");

            var path = System.IO.Path.Combine(AppContext.BaseDirectory, folder, file);
            if (System.IO.File.Exists(path))
            {
                try
                {
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.UriSource = new Uri(path, UriKind.Absolute);
                    bmp.EndInit();
                    bmp.Freeze();
                    info = $"OK file:{path}";
                    return bmp;
                }
                catch (Exception ex) { logs.AppendLine($"x file:{path} => {ex.Message}"); }
            }
            else logs.AppendLine($"x file not found:{path}");

            info = logs.ToString().TrimEnd();
            return null;
        }

        private static bool TryLoadBitmapFromPack(string packUri, out BitmapImage? bmp, out Exception? error)
        {
            bmp = null; error = null;
            try
            {
                var uri = new Uri(packUri, UriKind.Absolute);
                var sri = Application.GetResourceStream(uri);
                if (sri?.Stream == null) throw new IOException("GetResourceStream returned null");

                var b = new BitmapImage();
                b.BeginInit();
                b.CacheOption = BitmapCacheOption.OnLoad;
                b.StreamSource = sri.Stream;
                b.EndInit();
                b.Freeze();
                bmp = b;
                return true;
            }
            catch (Exception ex) { error = ex; return false; }
        }

        private MemoryStream? LoadStreamWithFallback(string folder, string file, out string info)
        {
            var logs = new System.Text.StringBuilder();

            if (TryLoadStreamFromPack($@"pack://application:,,,/;component/{folder}/{file}", out var ms1, out var e1))
            { info = $"OK component:{folder}/{file}"; return ms1; }
            logs.AppendLine($"x component: {e1?.Message}");

            if (TryLoadStreamFromPack($@"pack://application:,,,/{folder}/{file}", out var ms2, out var e2))
            { info = $"OK pack:{folder}/{file}"; return ms2; }
            logs.AppendLine($"x pack: {e2?.Message}");

            if (TryLoadFileStreamFromSiteOfOrigin($@"{folder}/{file}", out var ms3, out var e3))
            { info = $"OK siteoforigin:{folder}/{file}"; return ms3; }
            logs.AppendLine($"x siteoforigin: {e3?.Message}");

            var path = System.IO.Path.Combine(AppContext.BaseDirectory, folder, file);
            if (System.IO.File.Exists(path))
            {
                try
                {
                    var ms = new MemoryStream(System.IO.File.ReadAllBytes(path));
                    info = $"OK file:{path}";
                    return ms;
                }
                catch (Exception ex) { logs.AppendLine($"x file:{path} => {ex.Message}"); }
            }
            else logs.AppendLine($"x file not found:{path}");

            info = logs.ToString().TrimEnd();
            return null;
        }

        private static bool TryLoadStreamFromPack(string packUri, out MemoryStream? ms, out Exception? error)
        {
            ms = null; error = null;
            try
            {
                var uri = new Uri(packUri, UriKind.Absolute);
                var sri = Application.GetResourceStream(uri);
                if (sri?.Stream == null) throw new IOException("GetResourceStream returned null");

                var mem = new MemoryStream();
                sri.Stream.CopyTo(mem);
                mem.Position = 0;
                ms = mem;
                return true;
            }
            catch (Exception ex) { error = ex; return false; }
        }

        private static bool TryLoadFileStreamFromSiteOfOrigin(string relative, out MemoryStream? ms, out Exception? error)
        {
            ms = null; error = null;
            try
            {
                var uri = new Uri($@"pack://siteoforigin:,,,/{relative}", UriKind.Absolute);
                using var s = Application.GetRemoteStream(uri)?.Stream;
                if (s == null) throw new IOException("GetRemoteStream returned null");
                var mem = new MemoryStream();
                s.CopyTo(mem);
                mem.Position = 0;
                ms = mem;
                return true;
            }
            catch (Exception ex) { error = ex; return false; }
        }

        private static MemoryStream LoadMemoryStreamFromPack(string packUri)
        {
            var uri = new Uri(packUri, UriKind.Absolute);
            StreamResourceInfo? info = Application.GetResourceStream(uri);
            if (info?.Stream == null) throw new IOException($"リソースが見つかりません: {packUri}");

            var ms = new MemoryStream();
            info.Stream.CopyTo(ms);
            ms.Position = 0;
            return ms;
        }

        //private static BitmapImage LoadBitmapFromPack(string packUri)
        //{
        //    var uri = new Uri(packUri, UriKind.Absolute);
        //    StreamResourceInfo? info = Application.GetResourceStream(uri);
        //    if (info?.Stream == null) throw new IOException($"画像が見つかりません: {packUri}");

        //    var bmp = new BitmapImage();
        //    bmp.BeginInit();
        //    bmp.CacheOption = BitmapCacheOption.OnLoad;
        //    bmp.StreamSource = info.Stream;
        //    bmp.EndInit();
        //    bmp.Freeze();
        //    return bmp;
        //}

        // ===== 時計描画更新 =====
        private void UpdateClock()
        {
            var now = DateTime.Now;
            TxtNow.Text = now.ToString("HH:mm");

            double minuteAngle = now.Minute * 6.0 + now.Second * 0.1;
            double hourAngle = (now.Hour % 12) * 30.0 + now.Minute * 0.5 + now.Second * (0.5 / 60.0);

            RtMinute.Angle = minuteAngle;
            RtHour.Angle = hourAngle;
        }

        // ===== ドラッグで移動 =====
        private void RootGrid_MouseLeftButtonDown(object? sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                try { DragMove(); } catch { /* 無視 */ }
            }
        }

        // ===== アラーム関連 =====
        private void ApplyAlarmFromText()
        {
            string text = TxtAlarm.Text.Trim();
            if (TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out var ts))
            {
                _alarmTime = new TimeSpan(ts.Hours, ts.Minutes, 0);
                TxtStatus.Text = $"アラーム設定：{_alarmTime:hh\\:mm}";
            }
            else
            {
                _alarmTime = null;
                TxtStatus.Text = "アラーム時刻の形式が不正です（例：07:00 / 18:30）。";
            }
        }

        private void CheckAlarm()
        {
            if (!_alarmTime.HasValue) return;
            if (!ChkEnableAlarm.IsChecked.GetValueOrDefault()) return;
            if (_isRinging) return;

            var now = DateTime.Now;
            var nowHM = new TimeSpan(now.Hour, now.Minute, 0);
            if (nowHM == _alarmTime.Value)
            {
                StartAlarm();
            }
        }

        private void StartAlarm()
        {
            if (_player == null)
            {
                _wavStream = LoadMemoryStreamFromPack("pack://application:,,,/resources/Clock-Alarm04-01.wav");
                _player = new SoundPlayer(_wavStream);
                _player.Load();
            }

            try
            {
                _player.PlayLooping();
                _isRinging = true;
                BtnStop.IsEnabled = true;
                TxtStatus.Text = $"アラーム鳴動中：{_alarmTime:hh\\:mm}（停止ボタンで停止）";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "アラーム再生エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void StopAlarm()
        {
            if (_player == null) return;
            try
            {
                _player.Stop();
                _isRinging = false;
                BtnStop.IsEnabled = false;
                TxtStatus.Text = (_alarmTime.HasValue && ChkEnableAlarm.IsChecked == true)
                    ? $"アラーム待機中：{_alarmTime:hh\\:mm}"
                    : "アラーム未設定";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "アラーム停止エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ChkEnableAlarm_Checked(object sender, RoutedEventArgs e)
        {
            ApplyAlarmFromText();
            if (_alarmTime.HasValue)
                TxtStatus.Text = $"アラーム待機中：{_alarmTime:hh\\:mm}";
        }

        private void ChkEnableAlarm_Unchecked(object sender, RoutedEventArgs e)
        {
            StopAlarm();
            _alarmTime = null;
            TxtStatus.Text = "アラーム未設定";
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e) => StopAlarm();

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                _player?.Stop();
                _wavStream?.Dispose();
            }
            catch { /* 無視 */ }
            base.OnClosed(e);
        }
    }
}
