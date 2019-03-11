using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;

using System.Diagnostics;
using Tesseract;
using System.Text;
using Leptonica;


namespace TToyaCUP_OCR
{
    public partial class Form1 : Form
    {
        get_window frame;
       
        Rectangle global_rec;

        int monoImage_Threshold = 0;

        bool mouse_down = false;

        Color p1 = Color.Moccasin;
        Color p2 = Color.LightGreen;
        Color p3 = Color.Aquamarine;
        Color p4 = Color.HotPink;

        // penを作成
        Pen p1_pen;
        Pen p2_pen;
        Pen p3_pen;
        Pen p4_pen;

        int pen_size = 2;

        Dictionary<TextBox, TextBox> pear = new Dictionary<TextBox, TextBox>();

        #region DLL import
        [DllImport("User32.dll")]
        private extern static bool PrintWindow(IntPtr hwnd, IntPtr hDC, uint nFlags);

        [DllImport("User32.dll")]
        public extern static System.IntPtr GetDC(System.IntPtr hWnd);


        // 座標からウインドウのハンドルを取得する。
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT Point);

        // ハンドルからウインドウの位置を取得。
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern bool PostMessage(IntPtr hWnd, Int32 Msg, IntPtr wParam, IntPtr lParam);
        #endregion


        /// <summary>
        /// 目標ウィンドウハンドル用グローバル変数
        /// </summary>
        private IntPtr GlobalHWnd;

        /// <summary>
        /// 目標ウィンドウ用の範囲
        /// </summary>
        private RECT AimRect;

        public Form1()
        {
            InitializeComponent();
            pear[player1] = point1;
            pear[player2] = point2;
            pear[player3] = point3;
            pear[player4] = point4;
        }

        private void label1_MouseUp(object sender, MouseEventArgs e)
        {

        }

        /// <summary>
        /// マウスカーソルの下のウィンドウハンドルの収得
        /// </summary>
        private void Get_hWnd_Ander_Mouse()
        {
            //マウスカーソルの位置を取得（スクリーン座標）
            //GetCursorPos(lpt_Pos);
            POINT pos;
            RECT r;
            pos.x = Control.MousePosition.X;
            pos.y = Control.MousePosition.Y;
            Bitmap bmp, saveimg;

            //マウスカーソル下のウィンドウを取得
            GlobalHWnd = WindowFromPoint(pos);

            //指定ウィンドウの範囲ゲット
            GetWindowRect(GlobalHWnd, out r);

            int width = r.right - r.left;
            int height = r.bottom - r.top;

        }

        /// <summary>
        /// 画面全体のキャプチャ
        /// </summary>
        /// <returns></returns>
        public Bitmap CaptureAll()
        {

            SendKeys.SendWait("^{PRTSC}");//全画面をスクリーンキャプチャする場合

            IDataObject data = Clipboard.GetDataObject();

            Bitmap bmp = (Bitmap)data.GetData(DataFormats.Bitmap);

            return bmp;
        }

        /// <summary>
        /// 指定ハンドルをキャプチャ
        /// </summary>
        /// <param name="_hWnd">ウィンドウハンドル</param>
        /// <param name="_W">幅</param>
        /// <param name="_H">高さ</param>
        /// <returns></returns>
        public Bitmap CaptureControl(IntPtr _hWnd, int _W, int _H)
        {
            Bitmap img = new Bitmap(_W, _H);
            System.Drawing.Graphics memg = System.Drawing.Graphics.FromImage(img);
            IntPtr dc = memg.GetHdc();
            PrintWindow(_hWnd, dc, 0);
            memg.ReleaseHdc(dc);
            memg.Dispose();
            return img;
        }


        /// <summary>
        /// 余白の削除
        /// </summary>
        /// <param name="_bmp">削除したいビットマップ</param>
        /// <param name="_aimcolor">余白の色</param>
        /// <returns></returns>
        private Bitmap cut_Margin(Bitmap _bmp, Rectangle rectangle)
        {
            Bitmap re_bmp;

            //左の余白終了位置
            int left_margin = rectangle.X;

            //右の余白終了位置
            int right_margin = rectangle.X + rectangle.Width;

            //上の余白終了位置
            int top_margin = rectangle.Y;

            int Width = _bmp.Width;
            int Height = _bmp.Height;

            //return用bmp
            re_bmp = _bmp;

            //余白削除後の幅と高さ
            int W = rectangle.Width;
            int H = rectangle.Height;

            //debug 余白表示
            //textBox1.Text = "上余白" + top_margin.ToString() + "\r\n" + "左余白" + left_margin.ToString() + "\r\n" + "右余白" + right_margin.ToString();

            if (W > 0 && H > 0) //高さと幅が0以上なら
            {
                //byte[] byte_for_bmp = XyToByteAray(new Rectangle(0, 0, Width, Height), new Rectangle(top_margin, left_margin, W, H), bmp_base);
                //re_bmp = ByteToBmp(byte_for_bmp, new Rectangle(top_margin, left_margin, W, H));

                //余白削除したbmpをreturn用bmpに代入
                re_bmp = _bmp.Clone(rectangle, _bmp.PixelFormat);
            }

            return re_bmp;
        }

        /// <summary>
        /// 指定された画像から1bppのイメージを作成する
        /// </summary>
        /// <param name="img">基になる画像</param>
        /// <returns>1bppに変換されたイメージ</returns>
        public Bitmap Create1bppImage(Bitmap img)
        {
            //1bppイメージを作成する
            Bitmap newImg = new Bitmap(img.Width, img.Height,
                PixelFormat.Format1bppIndexed);

            //Bitmapをロックする
            BitmapData bmpDate = newImg.LockBits(new Rectangle(0, 0, newImg.Width, newImg.Height),ImageLockMode.WriteOnly, newImg.PixelFormat);

            //新しい画像のピクセルデータを作成する
            byte[] pixels = new byte[bmpDate.Stride * bmpDate.Height];
            for (int y = 0; y < bmpDate.Height; y++)
            {
                for (int x = 0; x < bmpDate.Width; x++)
                {
                    var R = (img.GetPixel(x, y).R / 255f ) * 0.2999;
                    var G = (img.GetPixel(x, y).G / 255f ) * 0.587;
                    var B = (img.GetPixel(x, y).B / 255f ) * 0.114;

                    //var V = img.GetPixel(x, y).GetBrightness();
                    var V = R+G+B;
                    
                    if (color_reversal.Checked)
                    {
                        if (monoImage_Threshold / 100f >= V)
                        {
                            //ピクセルデータの位置
                            int pos = (x >> 3) + bmpDate.Stride * y;
                            //白くする
                            pixels[pos] |= (byte)(0x80 >> (x & 0x7));
                        }
                    }
                    else
                    {
                        if (monoImage_Threshold / 100f <= V)
                        {
                            //ピクセルデータの位置
                            int pos = (x >> 3) + bmpDate.Stride * y;
                            //白くする
                            pixels[pos] |= (byte)(0x80 >> (x & 0x7));
                        }
                    }
                    
                }
            }
            //作成したピクセルデータをコピーする
            IntPtr ptr = bmpDate.Scan0;
            System.Runtime.InteropServices.Marshal.Copy(pixels, 0, ptr, pixels.Length);

            //ロックを解除する
            newImg.UnlockBits(bmpDate);

            return newImg;
        }

        /// <summary>
        /// 指定された画像からネガティブイメージを作成する
        /// </summary>
        /// <param name="img">基の画像</param>
        /// <returns>作成されたネガティブイメージ</returns>
        public static Image CreateNegativeImage(Image img)
        {
            //ネガティブイメージの描画先となるImageオブジェクトを作成
            Bitmap negaImg = new Bitmap(img.Width, img.Height);
            //negaImgのGraphicsオブジェクトを取得
            System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(negaImg);

            //ColorMatrixオブジェクトの作成
            System.Drawing.Imaging.ColorMatrix cm =
                new System.Drawing.Imaging.ColorMatrix();
            //ColorMatrixの行列の値を変更して、色が反転されるようにする
            cm.Matrix00 = -1;
            cm.Matrix11 = -1;
            cm.Matrix22 = -1;
            cm.Matrix33 = 1;
            cm.Matrix40 = cm.Matrix41 = cm.Matrix42 = cm.Matrix44 = 1;

            //ImageAttributesオブジェクトの作成
            System.Drawing.Imaging.ImageAttributes ia =
                new System.Drawing.Imaging.ImageAttributes();
            //ColorMatrixを設定する
            ia.SetColorMatrix(cm);

            //ImageAttributesを使用して色が反転した画像を描画
            g.DrawImage(img,
                new Rectangle(0, 0, img.Width, img.Height),
                0, 0, img.Width, img.Height, GraphicsUnit.Pixel, ia);

            //リソースを解放する
            g.Dispose();

            return negaImg;
        }


        private void label1_MouseDown(object sender, MouseEventArgs e)
        {
            Cursor.Current = Cursors.Cross;
        }

        private void player1_1_CheckedChanged(object sender, EventArgs e)
        {
            set_all();
        }

        private void set_all()
        {
            string str = "";

            str += player1.Text + "\t" + pear[player1].Text + "\r\n";
            str += player2.Text + "\t" + pear[player2].Text + "\r\n";
            str += player3.Text + "\t" + pear[player3].Text + "\r\n";
            str += player4.Text + "\t" + pear[player4].Text + "\r\n";


            string copy = "";

            copy += copy1.Checked ? pear[player1].Text + "\r\n" : "";
            copy += copy2.Checked ? pear[player2].Text + "\r\n" : "";
            copy += copy3.Checked ? pear[player3].Text + "\r\n" : "";
            copy += copy4.Checked ? pear[player4].Text + "\r\n" : "";

            Clipboard.SetData(DataFormats.Text, copy);

            label6.Text = copy;

            set_color();
        }

        private void set_color()
        {
            pear[player1].BackColor = p1;
            pear[player2].BackColor = p2;
            pear[player3].BackColor = p3;
            pear[player4].BackColor = p4;
            return;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            player1.BackColor = p1;
            player2.BackColor = p2;
            player3.BackColor = p3;
            player4.BackColor = p4;
            set_color();

            label6.Text = "";

            Properties.Settings.Default.Reload();

            trackBar1.Value = monoImage_Threshold = Properties.Settings.Default._Threshold;
            read_eng.Checked = Properties.Settings.Default._read_eng;

            color_reversal.Checked = Properties.Settings.Default._color_reversal;

            p1_pen = new Pen(p1, pen_size);
            p2_pen = new Pen(p2, pen_size);
            p3_pen = new Pen(p3, pen_size);
            p4_pen = new Pen(p4, pen_size);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

            Properties.Settings.Default._Threshold = monoImage_Threshold;
            Properties.Settings.Default._read_eng= read_eng.Checked;

            Properties.Settings.Default._color_reversal = color_reversal.Checked;

            Properties.Settings.Default.Save();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IDataObject data = Clipboard.GetDataObject();

            if (data.GetDataPresent(DataFormats.Text))
            {
                string[] lines = ((string)data.GetData(DataFormats.Text)).Split(new string[] { "\r\n" }, StringSplitOptions.None);
                if (lines.Length >= 4)
                {
                    player1.Text = lines[0];
                    player2.Text = lines[1];
                    player3.Text = lines[2];
                    player4.Text = lines[3];
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if(frame == null)
                frame = new get_window();

            frame.ShowInTaskbar = false; //タスクバーに表示させない

            frame.FrameBorderSize = 10; //線の太さ

            frame.FrameColor = Color.Blue; //線の色

            frame.AllowedTransform = true; //サイズ変更の可否

            frame.Show();

            set_rect_window.Enabled = false;
            ok.Enabled = true;
            get_screen.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (frame != null)
            {
                global_rec = frame.SelectedWindow;
                frame.Hide();

                set_rect_window.Enabled = true;
                ok.Enabled = false;
                get_screen.Enabled = true;

                int min_x = 0;
                int min_y = 0;
                //すべてのディスプレイを列挙する
                foreach (System.Windows.Forms.Screen s in System.Windows.Forms.Screen.AllScreens)
                {
                    min_x = s.Bounds.X < min_x ? s.Bounds.X : min_x;
                    min_y = s.Bounds.Y < min_y ? s.Bounds.Y : min_y;
                }

                global_rec.X = global_rec.X - min_x;
                global_rec.Y = global_rec.Y - min_y;
            }
        }

        private void get_screen_Click(object sender, EventArgs e)
        {
            //timer_hWnd.Stop();
            Get_hWnd_Ander_Mouse();
            Cursor.Current = Cursors.Default;
            GetWindowRect(GlobalHWnd, out AimRect);

            if (global_rec.Width == 0 && global_rec.Height == 0)
            {
                MessageBox.Show("取り込み範囲が指定されていません。");
                return;
                Rectangle r = new Rectangle() { X = 540, Y = 430, Width = 120, Height = 300 };
                pictureBox1.Image = Create1bppImage(cut_Margin(CaptureControl(GlobalHWnd, 680, 770), r));
            }
            else
            {
                pictureBox1.Image = Create1bppImage(resize_img(cut_Margin(CaptureAll(), global_rec)));
                //pictureBox1.Image = cut_Margin(CaptureAll(), global_rec);
            }

            string path = AppDomain.CurrentDomain.BaseDirectory;

            Bitmap bmp = color_reversal.Checked ? (Bitmap)pictureBox1.Image : (Bitmap)CreateNegativeImage(pictureBox1.Image);

            bmp.Save(path + "_tmp.bmp");
            var lines =  run_tessract();

            if (lines.Count == 0) return;

            string[] nums = new string[4];
            int count = 0;
            textBox1.Text = string.Join("\r\n", lines.ToArray());
            foreach (string item in lines)
            {
                var str = Regex.Replace(item, @"\s", "");
                if (item.Length > 0)
                {
                    nums[count] = str;
                    count++;
                }

                if (count >= 4)
                    break;
            }
            point1.Text = nums[0];
            point2.Text = nums[1];
            point3.Text = nums[2];
            point4.Text = nums[3];

            set_all();

        }
        /*
        private List<string> run_tessract()
        {
            List<string> re = new List<string>();
            try
            {
                string lang = 1 == 1 ? "eng" : "jpn";
                using (var engine = new TesseractEngine(@"./tessdata", lang, EngineMode.TesseractOnly))
                {
                    if (!read_eng.Checked)
                        engine.SetDebugVariable("tessedit_char_whitelist", "1234567890");
                    using (var img = Pix.LoadFromFile(@"./_tmp.bmp"))
                    {
                        using (var page = engine.Process(img))
                        {
                            var text = page.GetText();
                            Console.WriteLine("Mean confidence: {0}", page.GetMeanConfidence());

                            Console.WriteLine("Text (GetText): \r\n{0}", text);
                            Console.WriteLine("Text (iterator):");
                            using (var iter = page.GetIterator())
                            {
                                iter.Begin();
                                do
                                {
                                    do
                                    {
                                        do
                                        {
                                            string str = "";
                                            do
                                            {
                                                if (iter.IsAtBeginningOf(PageIteratorLevel.Block))
                                                {
                                                    //re += "\r\n";
                                                }

                                                str += (iter.GetText(PageIteratorLevel.Word));
                                                str += (" ");

                                                if (iter.IsAtFinalOf(PageIteratorLevel.TextLine, PageIteratorLevel.Word))
                                                {
                                                    str += " ";
                                                }
                                            } while (iter.Next(PageIteratorLevel.TextLine, PageIteratorLevel.Word));

                                            if (iter.IsAtFinalOf(PageIteratorLevel.Para, PageIteratorLevel.TextLine))
                                            {
                                                str += " ";
                                            }

                                            re.Add(str);
                                        } while (iter.Next(PageIteratorLevel.Para, PageIteratorLevel.TextLine));
                                    } while (iter.Next(PageIteratorLevel.Block, PageIteratorLevel.Para));
                                } while (iter.Next(PageIteratorLevel.Block));
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Trace.TraceError(e.ToString());
                textBox1.Text += ("Unexpected Error: " + e.Message);
                textBox1.Text += ("Details: ");
                textBox1.Text += (e.ToString());

            }

            return re;
        }
        */

        private Bitmap resize_img(Bitmap _in_bmp)
        {
            float ratio = 2f;
            int W = (int)Math.Ceiling(_in_bmp.Width * ratio);
            int H = (int)Math.Ceiling(_in_bmp.Height * ratio);

            //描画先とするImageオブジェクトを作成する
            Bitmap canvas = new Bitmap(W, H);
            //ImageオブジェクトのGraphicsオブジェクトを作成する
            System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(canvas);

            //Bitmapオブジェクトの作成
            Bitmap image = (Bitmap)_in_bmp.Clone();

            //補間方法として高品質双三次補間を指定する
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            g.DrawImage(image, 0, 0, _in_bmp.Width * ratio, _in_bmp.Height * ratio);

            //BitmapとGraphicsオブジェクトを破棄
            image.Dispose();
            g.Dispose();

            //PictureBox1に表示する
            return canvas;
        }

        private List<string> run_tessract()
        {
            List<string> re = new List<string>();
            string dataPath = "./tessdata/";
            string language = 1 != 1 ? "eng" : "jpn";
            string inputFile = "./_tmp.bmp";
            OcrEngineMode oem = OcrEngineMode.DEFAULT;
            PageSegmentationMode psm = PageSegmentationMode.AUTO_OSD;

            TessBaseAPI tessBaseAPI = new TessBaseAPI();

            // Initialize tesseract-ocr 
            if (!tessBaseAPI.Init(dataPath, language, oem))
            {
                throw new Exception("Could not initialize tesseract.");
            }

            // Set the Page Segmentation mode
            tessBaseAPI.SetPageSegMode(psm);

            // Set the input image
            Pix pix = tessBaseAPI.SetImage(inputFile);

            tessBaseAPI.SetVariable("number","1234567890");

            // Recognize image
            tessBaseAPI.Recognize();

            ResultIterator resultIterator = tessBaseAPI.GetIterator();

            // extract text from result iterator
            StringBuilder stringBuilder = new StringBuilder();
            PageIteratorLevel pageIteratorLevel = PageIteratorLevel.RIL_PARA;
            do
            {
                string str = resultIterator.GetUTF8Text(pageIteratorLevel);
                               

                if(str != null)
                {
                    str = Regex.Replace(str, @"\n", "\r\n");
                    re.Add(str);
                }
                    
            } while (resultIterator.Next(pageIteratorLevel));

            tessBaseAPI.Dispose();
            pix.Dispose();
            return re;
        }

        private void ライセンスToolStripMenuItem_Click(object sender, EventArgs e)
        {
            License lic = new License();

            lic.ShowDialog(this);
            lic.Dispose();
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            label7.Text = trackBar1.Value.ToString();
            monoImage_Threshold = trackBar1.Value;
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            if (!mouse_down)
            {
                // Graphicsオブジェクトの作成
                System.Drawing.Graphics g = this.CreateGraphics();
                
                Point[] p1_lines = get_line(player1, pear[player1]);
                Point[] p2_lines = get_line(player2, pear[player2]);
                Point[] p3_lines = get_line(player3, pear[player3]);
                Point[] p4_lines = get_line(player4, pear[player4]);

                // lineを描画
                drow_edge_lines(g, p1_pen, p1_lines);
                drow_edge_lines(g, p2_pen, p2_lines);
                drow_edge_lines(g, p3_pen, p3_lines);
                drow_edge_lines(g, p4_pen, p4_lines);

                // Graphicsを解放する
                g.Dispose();
            }
        }

        private void drow_edge_lines(System.Drawing.Graphics g, Pen _p, Point[] lines)
        {
            Pen black_pen = new Pen(Color.Black, _p.Width + 2);

            // lineを描画
            g.DrawLines(black_pen, lines);
            g.DrawLines(_p, lines);
        }

        private Point[] get_line(TextBox _start, TextBox _end)
        {
            int X = _start.Location.X + _start.Size.Width;
            int Y = _start.Location.Y + (_start.Size.Height / 2);

            Point start = new Point(X,Y);

            X = _end.Location.X;
            Y = _end.Location.Y + (_start.Size.Height / 2);

            Point end = new Point(X, Y);
            Point[] re = { start, end };

            return re;
        }

        private Point[] get_line(TextBox _start)
        {
            int X = _start.Location.X + _start.Size.Width;
            int Y = _start.Location.Y + (_start.Size.Height / 2);

            Point start = new Point(X, Y);

            //フォーム上の座標でマウスポインタの位置を取得する
            //画面座標でマウスポインタの位置を取得する
            System.Drawing.Point sp = System.Windows.Forms.Cursor.Position;
            //画面座標をクライアント座標に変換する
            System.Drawing.Point cp = this.PointToClient(sp);
            //X座標を取得する
            X = cp.X;
            //Y座標を取得する
            Y = cp.Y;
            
            Point end = new Point(X, Y);
            Point[] re = { start, end };

            return re;
        }


        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(Point point);

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void player1_MouseDown(object sender, MouseEventArgs e)
        {
            Cursor.Current = Cursors.Hand;
            mouse_down = true;
        }

        private void player1_MouseMove(object sender, MouseEventArgs e)
        {
            // Graphicsオブジェクトの作成
            System.Drawing.Graphics g = this.CreateGraphics();
            
            // lineの始点と終点を設定

            switch (((TextBox)sender).Name)
            {
                case "player1":
                    drow_edge_lines(g, p1_pen, get_line(player1));
                    break;

                case "player2":
                    drow_edge_lines(g, p2_pen, get_line(player2));
                    break;

                case "player3":
                    drow_edge_lines(g, p3_pen, get_line(player3));
                    break;

                case "player4":
                    drow_edge_lines(g, p4_pen, get_line(player4));
                    break;
            }
            
            // Graphicsを解放する
            g.Dispose();
            this.Invalidate();
        }

        private void player1_MouseUp(object sender, MouseEventArgs e)
        {
            Cursor.Current = Cursors.Default;
            IntPtr handle = WindowFromPoint(Control.MousePosition);
            if (handle != IntPtr.Zero)
            {
                Control control = Control.FromHandle(handle);
                if (control != null)
                {
                    switch (control.Name)
                    {
                        case "point1":
                            pear[(TextBox)sender] = point1;
                            break;

                        case "point2":
                            pear[(TextBox)sender] = point2;
                            break;

                        case "point3":
                            pear[(TextBox)sender] = point3;
                            break;

                        case "point4":
                            pear[(TextBox)sender] = point4;
                            break;
                    }

                    mouse_down = false;
                    this.Invalidate();
                    set_all();

                }
            }
        }
    }
}
