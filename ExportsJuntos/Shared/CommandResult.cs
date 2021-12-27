using Flunt.Notifications;

namespace ExportsJuntos.Shared
{
    public class CommandResult : Notifiable<Notification>
    {
        public object Response { get; set; }

        public void Merge(CommandResult from)
        {
            AddNotifications(from.Notifications);
            Response = from.Response;
        }

        public virtual bool Validate() => IsValid;
    }
}
