using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Cors;
using InventorApp.API.Models;
using InventorApp.API.Services;

namespace InventorApp.API.Controllers
{
    [Route("api/auth")]
    [ApiController]
    [EnableCors("AllowSpecificOrigin")]
    public class AuthController : ControllerBase
    {
        private readonly UserService _userService;

        public AuthController(UserService userService)
        {
            _userService = userService;
        }

        [HttpPost("login")]
        public async Task<ActionResult<LoginResponse>> Login([FromBody] LoginRequest request)
        {
            var result = await _userService.LoginAsync(request);
            if (!result.Success) return Unauthorized(result);
            return Ok(result);
        }

        [HttpPost("forgot-password")]
        public async Task<ActionResult<ForgotPasswordResponse>> ForgotPassword([FromBody] ForgotPasswordRequest request)
        {
            var result = await _userService.ChangePasswordAsync(request);
            if (!result.Success) return BadRequest(result);
            return Ok(result);
        }
    }
}

