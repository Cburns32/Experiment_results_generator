using System.Windows.Forms;

namespace System
{
    internal class MouseEventHandler
    {
        private Action<object, MouseEventArgs> fileLabel1_MouseDown;

        public MouseEventHandler(Action<object, MouseEventArgs> fileLabel1_MouseDown)
        {
            this.fileLabel1_MouseDown = fileLabel1_MouseDown;
        }
    }
}