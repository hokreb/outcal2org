using System;
using Microsoft.Office.Interop.Outlook;


namespace outcal2org
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 3)
            {
                GetOutlookCalendarItems(int.Parse(args[1]), int.Parse(args[2]));
            }
            else
            {
                GetOutlookCalendarItems();
            }
        }

        public static void GetOutlookCalendarItems(int start = -30, int stop = 90)
        {
            Application app = new Application();
            NameSpace mapiNamespace = app.GetNamespace("MAPI"); ;
            MAPIFolder CalendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            Items calendarItems = CalendarFolder.Items;

            calendarItems.IncludeRecurrences = true;
            calendarItems.Sort("[Start]");

            var actDate = DateTime.Now;

            var begin = actDate + new TimeSpan(start, 0, 0, 0);
            var end = actDate + new TimeSpan(stop, 0, 0, 0);

            var restrict = $"[Start] >= '{begin.ToString("dd/MM/yyyy")}' AND [END] <= '{end.ToString("dd/MM/yyyy")}'";                        

            var restrictedItems = calendarItems.Restrict(restrict);

            Console.WriteLine("* Outlook Kalender (by outcal2org)");
            foreach (AppointmentItem item in restrictedItems)
            {
                Console.WriteLine($"** {item.Subject}");

                if (item.AllDayEvent)
                {
                    if (item.Duration < (24 * 60))
                    {
                        Console.WriteLine($"<{item.Start.ToString("yyyy-MM-dd")}>");
                    }
                    else
                    {
                        Console.WriteLine($"<{item.Start.ToString("yyyy-MM-dd")}>-<{item.End.ToString("yyyy-MM-dd")}>");
                    }
                }
                else
                {
                    Console.WriteLine($"<{item.Start.ToString("yyyy-MM-dd HH:mm")}-{item.End.ToString("HH:mm")}>" );
                }

                Console.WriteLine();
            }
        }
    }
}
