using Microsoft.AspNetCore.Mvc;
using Task3.Helper;

namespace Task3.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CalendarController : ControllerBase
    {
        [HttpPost]
        public async Task<IActionResult> CreateGoogleCalendar([FromBody] Models.Calendar request)
        {
            return Ok(await CalendarHelper.CreateGoogleCalendar(request));
        }

        [HttpGet]
        public async Task<IActionResult> GetUpcomingEvents(string? searchText)
        {
            return Ok(await CalendarHelper.GetUpcomingEvents(searchText));
        }

        [HttpDelete]
        public async Task DeleteGoogleCalendar(string eventId)
        {
            await CalendarHelper.DeleteGoogleCalendar(eventId);
        }

        [HttpPatch]
        public async Task EditGoogleCalendar(string eventId,[FromBody] Models.Calendar request)
        {
            await CalendarHelper.EditGoogleCalendar(eventId, request);
        }

        [HttpPut]
        public async Task SetEventReminders(string eventId, int minutesBeforeStart)
        {
            await CalendarHelper.SetEventReminders(eventId, minutesBeforeStart);
        }

    }
}
