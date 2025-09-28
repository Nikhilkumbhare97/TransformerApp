using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Cors;
using InventorApp.API.Models;
using InventorApp.API.Services;

namespace InventorApp.API.Controllers
{
    [Route("api/users")]
    [ApiController]
    [EnableCors("AllowSpecificOrigin")]
    public class UserController : ControllerBase
    {
        private readonly UserService _userService;

        public UserController(UserService userService)
        {
            _userService = userService;
        }

        [HttpPost]
        public async Task<ActionResult<User>> CreateUser([FromBody] User userInput)
        {
            var created = await _userService.CreateUser(userInput);
            return CreatedAtAction(nameof(GetUserById), new { userUniqueId = created.UserUniqueId }, created);
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<User>>> GetAllUsers()
        {
            var users = await _userService.GetAllUsers();
            return Ok(users);
        }

        [HttpGet("{userUniqueId}")]
        public async Task<ActionResult<User>> GetUserById(long userUniqueId)
        {
            var user = await _userService.GetUserById(userUniqueId);
            if (user == null) return NotFound();
            return Ok(user);
        }

        [HttpGet("by-employee/{employeeId}")]
        public async Task<ActionResult<User>> GetByEmployeeId(string employeeId)
        {
            var user = await _userService.GetByEmployeeId(employeeId);
            if (user == null) return NotFound();
            return Ok(user);
        }

        [HttpPut("{userUniqueId}")]
        public async Task<ActionResult<User>> UpdateUser(long userUniqueId, [FromBody] UserUpdate partial)
        {
            var toPartial = new User
            {
                Name = partial.Name ?? string.Empty,
                EmployeeId = partial.EmployeeId ?? string.Empty,
                Email = partial.Email ?? string.Empty,
                PasswordEncrypted = partial.Password ?? string.Empty,
                Role = partial.Role ?? string.Empty,
                Status = partial.Status ?? string.Empty
            };
            var updated = await _userService.UpdateUser(userUniqueId, toPartial);
            if (updated == null) return NotFound();
            return Ok(updated);
        }

        [HttpDelete("{userUniqueId}")]
        public async Task<IActionResult> DeleteUser(long userUniqueId)
        {
            await _userService.DeleteUser(userUniqueId);
            return NoContent();
        }
    }
}

