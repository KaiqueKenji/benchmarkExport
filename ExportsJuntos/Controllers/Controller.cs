using ExportsJuntos.Shared;
using Flunt.Notifications;
using Microsoft.AspNetCore.Mvc;
using System.Net;
using Mvc = Microsoft.AspNetCore.Mvc;

namespace ExportsJuntos.Controllers;
public class Controller : Mvc.Controller
{
    protected new IActionResult Response(CommandResult commandResult, HttpStatusCode httpStatusCode = HttpStatusCode.OK) =>
        Response(commandResult, (int)httpStatusCode);

    protected new IActionResult Response(CommandResult commandResult, int statusCode)
    {
        if (commandResult.Notifications.Any()) return BadRequest(commandResult.Notifications);
        return StatusCode(statusCode, commandResult.Response);
    }

    protected IActionResult Ok<T>(IEnumerable<T> list)
    {
        if (list == null || !list.Any()) return NoContent();
        return Ok(list as object);
    }

    protected IActionResult BadRequest(string property, string message) =>
        BadRequest(new Notification[] { new(property, message) });
}

