namespace InventorApp.API.Models
{
    public class LoginRequest
    {
        public string Email { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
    }

    public class LoginResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public long? UserUniqueId { get; set; }
        public string? Name { get; set; }
        public string? EmployeeId { get; set; }
        public string? Email { get; set; }
        public string? Role { get; set; }
        public string? Status { get; set; }
    }

    public class ForgotPasswordRequest
    {
        public string Email { get; set; } = string.Empty;
        public string EmployeeId { get; set; } = string.Empty;
        public string OldPassword { get; set; } = string.Empty;
        public string NewPassword { get; set; } = string.Empty;
    }

    public class ForgotPasswordResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}

