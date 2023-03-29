using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.PeopleService.v1;
using MimeKit;
using MailKit.Net.Smtp;
using Google.Apis.Gmail.v1;
using Google.Apis.Util;
using Task3.Models;

namespace Task3.Helper
{
    public class CalendarHelper
    {
        public static async Task<Event?> CreateGoogleCalendar(Models.Calendar request)
        {
            if (request.Start >= request.End)
            {
                throw new ArgumentException("End time must be after start time");
            }
            try
            {
                string[] scopes = { CalendarService.Scope.CalendarEvents, PeopleServiceService.Scope.UserinfoEmail, PeopleServiceService.Scope.UserinfoProfile, GmailService.Scope.GmailCompose, GmailService.Scope.GmailSend };
                string ApplicationName = "Google Calendar Api";
                UserCredential credential;
                using (var stream = new FileStream(Path.Combine(Directory.GetCurrentDirectory(), "credential", "credential.json"), FileMode.Open, FileAccess.Read))
                {
                    string credPath = "token.json";
                    credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true));
                }

                if (credential.Token.IsExpired(SystemClock.Default))
                {
                    if (await credential.RefreshTokenAsync(CancellationToken.None))
                    {
                        var store = new FileDataStore("token.json", true);
                        store.StoreAsync(credential.UserId, credential).Wait();
                    }
                    else
                    {
                        throw new Exception("Failed to refresh token");
                    }
                }
                var services = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });

                var peopleService = new PeopleServiceService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });

                Event eventCalendar = new()
                {
                    Summary = request.Summary,
                    Location = request.Location,
                    Start = new EventDateTime
                    {
                        DateTime = request.Start,
                        TimeZone = "Africa/Cairo",
                    },
                    End = new EventDateTime
                    {
                        DateTime = request.End,
                        TimeZone = "Africa/Cairo",
                    },
                    Description = request.Description,
                };

                var eventRequest = services.Events.Insert(eventCalendar, "primary");
                var requestCreate = await eventRequest.ExecuteAsync();

                //--------------
                var person = peopleService.People.Get("people/me");
                person.PersonFields = "emailAddresses";
                var userEmail = person.Execute().EmailAddresses.FirstOrDefault()?.Value;
                Console.Out.WriteLine("Email____________________________________________________________________________>>>>>>>" + userEmail);
                if (!string.IsNullOrEmpty(userEmail))
                {
                    var message = new MimeMessage();
                    message.From.Add(new MailboxAddress("mmmm mmmm", "examtes112@outlook.com"));
                    message.To.Add(new MailboxAddress("Mohamed Naguib", userEmail));
                    message.Subject = "New Calendar Event Created";
                    message.Body = new TextPart("plain")
                    {
                        Text = $"A new calendar event has been created with the following details:\n\n" +
                        $"Summary: {request.Summary}\n" +
                        $"Location: {request.Location}\n" +
                        $"Start Time: {request.Start}\n" +
                        $"End Time: {request.End}\n" +
                        $"Description: {request.Description}\n"
                    };

                    using var client = new SmtpClient();
                    client.Connect("smtp.office365.com", 587, MailKit.Security.SecureSocketOptions.StartTlsWhenAvailable);
                    client.Authenticate("examtes112@outlook.com", "Assignment123");
                    client.Send(message);
                    client.Disconnect(true);

                }
                //-----------------
                return requestCreate;
            }
            catch (ArgumentNullException ex)            // Handle invalid input

            {
                Console.WriteLine($"Invalid argument: {ex.Message}");
                return null;
            }
            catch (AggregateException ex)              // Handle authentication errors
            {
                Console.WriteLine($"Authentication error: {ex.Message}");
                return null;
            }
            
            catch (Exception ex)                       // Handle other errors
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }

     
        public static Task<List<string>> GetUpcomingEvents(string? searchText)
        {
            var service = Authentication.Authenticate();
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // If Input empty retrieve all events, else search by event title, date, time, and description.
            if (!string.IsNullOrEmpty(searchText))
            {
                request.Q = searchText;
            }

            Events events = request.Execute();                  // Retrieve the list of events.

            List<string> upcomingEvents = new();

            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string? startTime = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(startTime))
                    {
                        startTime = eventItem.Start.Date;
                    }
                    upcomingEvents.Add($"ID: {eventItem.Id}, Summary: {eventItem.Summary}, Description: {eventItem.Description}, StartTime: {startTime}, EndTime: {eventItem.End.DateTime}");
                }
            }
            else
            {
                upcomingEvents.Add("No upcoming events found.");
            }

            return Task.FromResult(upcomingEvents);
        }

        
        public static async Task<Event> EditGoogleCalendar(string eventId, Models.Calendar request)
        {
            var service = Authentication.Authenticate();
            EventsResource.GetRequest getRequest = service.Events.Get("primary", eventId);      //Get the existing event by its ID
            Event existingEvent = await getRequest.ExecuteAsync();
            existingEvent.Summary = request.Summary;                                            //Update the event with the new information
            existingEvent.Location = request.Location;
            existingEvent.Description = request.Description;
            existingEvent.Start.DateTime = request.Start;
            existingEvent.End.DateTime = request.End;
            EventsResource.UpdateRequest updateRequest = service.Events.Update(existingEvent, "primary", eventId);
            Event updatedEvent = await updateRequest.ExecuteAsync();
            return updatedEvent;
        }

        
        public static async Task DeleteGoogleCalendar(string eventId)
        {
            var service = Authentication.Authenticate();
            EventsResource.ListRequest request = service.Events.List("primary");
            var eventRequest = service.Events.Delete("primary", eventId);                     //Delete the existing event by its ID
            await eventRequest.ExecuteAsync();
        }

        
        public static async Task SetEventReminders(string eventId, int minutesBeforeStart)
        {
            var service = Authentication.Authenticate();
            EventsResource.ListRequest request = service.Events.List("primary");
            var existingEventRequest = service.Events.Get("primary", eventId);                 //Get the existing event by its ID
            var existingEvent = await existingEventRequest.ExecuteAsync();
            var reminder = new EventReminder()
            {
                Minutes = minutesBeforeStart,
                Method = "popup"
            };
            existingEvent.Reminders = new Event.RemindersData()
            {
                UseDefault = false,
                Overrides = new List<EventReminder>() { reminder }
            };
            var updateRequest = service.Events.Update(existingEvent, "primary", eventId);       // Update the event with the new reminder
            await updateRequest.ExecuteAsync();
        }
    }
}