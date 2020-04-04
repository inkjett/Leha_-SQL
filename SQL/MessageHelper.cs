using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQL
{ // класс генерации сообщений, используются патерны, генерируются один раз
    public class MessageHelper 
    {
        private static readonly Lazy<MessageHelper> _instance = new Lazy<MessageHelper>(() => new MessageHelper());
        private MessageHelper()
        {

        }
        public delegate void MessageHandler(object sender, MessageGenerateEventArgs e);
        public event MessageHandler MessageGeneratedEventHandler;

        public static MessageHelper GetInstance() => _instance.Value;

        public void SetMessage(string message)
        {
            MessageGeneratedEventHandler?.Invoke(this, new MessageGenerateEventArgs(message));
        }
    }

    public class MessageGenerateEventArgs : EventArgs
    {
        public string Message { get; }
        public MessageGenerateEventArgs(string message)
        {
            Message = message;
        }
    }
   
}
