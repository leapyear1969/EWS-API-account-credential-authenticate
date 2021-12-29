using System;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;


namespace EWS_API_account_credential_authenticate
{
	class Program
	{
            static async System.Threading.Tasks.Task Main(string[] args)
            {

                // Create the binding.
                ExchangeService service = new ExchangeService();
                // Set the credentials for the on-premises server.
                service.Credentials = new WebCredentials("yourAccountHere", "yourPWDHere");
            // Set the URL.
            service.Url = new Uri("https://partner.outlook.cn/EWS/Exchange.asmx");

                //bind a user
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "someone@contosoc.com");
            // Set the URL.
            //service.Url = new Uri("https://computername.domain.contoso.com/EWS/Exchange.asmx");

            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://partner.outlook.cn/EWS.AccessAsUser.All" };
                //var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };

                // Make a  EWS call
                // Initialize values for the start and end times, and the number of appointments to retrieve.
                DateTime startDate = DateTime.Now;
                DateTime endDate = startDate.AddDays(30);
                const int NUM_APPTS = 5;
                // Initialize the calendar folder object with only the folder ID. 
                CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
                // Set the start and end time and number of appointments to retrieve.
                CalendarView cView = new CalendarView(startDate, endDate, NUM_APPTS);
                // Limit the properties returned to the appointment's subject, start time, and end time.
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);
                // Retrieve a collection of appointments by using the calendar view.
                FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

                Console.WriteLine("\nThe first " + NUM_APPTS + " appointments on your calendar from " + startDate.Date.ToShortDateString() +
                                  " to " + endDate.Date.ToShortDateString() + " are: \n");
                ItemId appointmentId = appointments.Items[0].Id;

                //foreach (Appointment a in appointments)
                //{
                //    //Console.Write("Subject: " + a.Subject.ToString() + " ");
                //    //Console.Write("Start: " + a.Start.ToString() + " ");
                //    //Console.Write("End: " + a.End.ToString());
                //    //Console.Write("appointmentId:" + a.Id.ToString());
                //    //Console.WriteLine();
                //    appointmentId = a.Id;
                //}


                // Instantiate an appointment object by binding to it by using the ItemId.
                // As a best practice, limit the properties returned to only the ones you need.
                Appointment appointment = Appointment.Bind(service, appointmentId, new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End));
                string oldSubject = appointment.Subject;
                appointment.Load();
                //Console.WriteLine(appointment.GetLoadedPropertyDefinitions());
                // Update properties on the appointment with a new subject, start time, and end time.
                appointment.Subject = appointment.Subject + "YourSubjects";
                appointment.Start.AddHours(25);
                appointment.End.AddHours(25);
                // Unless explicitly specified, the default is to use SendToAllAndSaveCopy.
                // This can convert an appointment into a meeting. To avoid this,
                // explicitly set SendToNone on non-meetings.
                SendInvitationsOrCancellationsMode mode = appointment.IsMeeting ? SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy : SendInvitationsOrCancellationsMode.SendToNone;
                // Send the update request to the Exchange server.
                appointment.Update(ConflictResolutionMode.AlwaysOverwrite, mode);
                // Verify the update.
                Console.WriteLine("Subject for the appointment was \"" + oldSubject + "\". The new subject is \"" + appointment.Subject + "\"");
            }
        }

    }


