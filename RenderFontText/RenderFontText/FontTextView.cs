using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RenderFontText
{
    public partial class FontTextView : Form
    {

        private ArrayList display_text;
        private Font display_font;
        private Brush display_brush;
        private int line_height;
        private static int margin = 10;

        public FontTextView(string font_name)
        {
            InitializeComponent();
            display_text = new ArrayList();
            display_font = new Font(font_name, 12);
            line_height = display_font.Height;
            display_brush = Brushes.Blue;
        }

        private int world_y_to_line_indx(int y)
        {
            if (y < margin)
                return -1;
            return (int)((y - margin) / line_height);
        }

        private Point line_indx_to_world(int indx)
        {
            Point top_left = new Point((int)margin, (int)((line_height * indx) + margin));
            return top_left;
        }


        public void render_text(string text)
        {
           // get graphics context for form

            display_text.Add(text);
            this.Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            // get graphics context
            Graphics dc = e.Graphics;
            int scroll_pos_x = this.AutoScrollPosition.X;
            int scroll_pos_y = this.AutoScrollPosition.Y;
            dc.TranslateTransform(scroll_pos_x, scroll_pos_y);

            if (display_text.Count > 0)
            {
                // we have lines of text to display
                // determine which lines are in the clipping region

                int min_line_num = world_y_to_line_indx(e.ClipRectangle.Top - scroll_pos_y);
                if (min_line_num == -1)
                    min_line_num = 0;

                int max_line_num = world_y_to_line_indx(e.ClipRectangle.Bottom - scroll_pos_y);
                if ((max_line_num == -1)||(max_line_num > (display_text.Count-1)))
                    max_line_num = display_text.Count - 1;

                string nxt_line;
                for (int indx = min_line_num; indx <= max_line_num; indx++)
                {
                    nxt_line = (string)display_text[indx];
                    dc.DrawString(nxt_line, display_font, display_brush, line_indx_to_world(indx));
                }
            }


        }
    }
}
