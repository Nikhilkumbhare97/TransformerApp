using System.Collections.Generic;
using System.Threading.Tasks;
using InventorApp.API.Models;
using InventorApp.API.Repositories;

namespace InventorApp.API.Services
{
    public class UserService
    {
        private readonly IUserRepository _userRepository;
        private readonly EncryptionService _encryptionService;

        public UserService(IUserRepository userRepository, EncryptionService encryptionService)
        {
            _userRepository = userRepository;
            _encryptionService = encryptionService;
        }

        public async Task<User> CreateUser(User userWithPlainPassword)
        {
            // Always encrypt the incoming value (bound from JSON as "password") before storing
            var encrypted = _encryptionService.EncryptToBase64(userWithPlainPassword.PasswordEncrypted);
            userWithPlainPassword.PasswordEncrypted = encrypted;
            return await _userRepository.CreateAsync(userWithPlainPassword);
        }

        public async Task<User?> GetUserById(long userUniqueId, bool includeDecryptedPassword = false)
        {
            var user = await _userRepository.GetByIdAsync(userUniqueId);
            if (user == null) return null;
            // Never decrypt for GET responses
            return user;
        }

        public async Task<IEnumerable<User>> GetAllUsers(bool includeDecryptedPassword = false)
        {
            var users = await _userRepository.GetAllAsync();
            return users;
        }

        public async Task<User?> GetByEmployeeId(string employeeId, bool includeDecryptedPassword = false)
        {
            var user = await _userRepository.GetByEmployeeIdAsync(employeeId);
            if (user == null) return null;
            // Never decrypt for GET responses
            return user;
        }

        public async Task<LoginResponse> LoginAsync(LoginRequest request)
        {
            var user = await _userRepository.GetByEmailAsync(request.Email);
            if (user == null)
            {
                return new LoginResponse { Success = false, Message = "User not found" };
            }
            var encryptedInput = _encryptionService.EncryptToBase64(request.Password);
            if (user.PasswordEncrypted != encryptedInput)
            {
                return new LoginResponse { Success = false, Message = "Invalid password" };
            }
            return new LoginResponse
            {
                Success = true,
                Message = "Login successful",
                UserUniqueId = user.UserUniqueId,
                Name = user.Name,
                EmployeeId = user.EmployeeId,
                Email = user.Email,
                Role = user.Role,
                Status = user.Status
            };
        }

        public async Task<ForgotPasswordResponse> ChangePasswordAsync(ForgotPasswordRequest request)
        {
            var byEmail = await _userRepository.GetByEmailAsync(request.Email);
            if (byEmail == null)
            {
                return new ForgotPasswordResponse { Success = false, Message = "User with given email does not exist" };
            }

            if (!string.Equals(byEmail.EmployeeId, request.EmployeeId))
            {
                return new ForgotPasswordResponse { Success = false, Message = "EmployeeId does not match the email" };
            }

            var encOld = _encryptionService.EncryptToBase64(request.OldPassword);
            if (!string.Equals(byEmail.PasswordEncrypted, encOld))
            {
                return new ForgotPasswordResponse { Success = false, Message = "Old password is incorrect" };
            }

            byEmail.PasswordEncrypted = _encryptionService.EncryptToBase64(request.NewPassword);
            await _userRepository.UpdateAsync(byEmail);
            return new ForgotPasswordResponse { Success = true, Message = "Password updated successfully" };
        }

        public async Task<User?> UpdateUser(long userUniqueId, User partial)
        {
            var existing = await _userRepository.GetByIdAsync(userUniqueId);
            if (existing == null) return null;

            if (!string.IsNullOrEmpty(partial.Name)) existing.Name = partial.Name;
            if (!string.IsNullOrEmpty(partial.EmployeeId)) existing.EmployeeId = partial.EmployeeId;
            if (!string.IsNullOrEmpty(partial.Email)) existing.Email = partial.Email;
            if (!string.IsNullOrEmpty(partial.Role)) existing.Role = partial.Role;
            if (!string.IsNullOrEmpty(partial.Status)) existing.Status = partial.Status;

            // If password provided, encrypt before saving
            if (!string.IsNullOrEmpty(partial.PasswordEncrypted))
            {
                existing.PasswordEncrypted = _encryptionService.EncryptToBase64(partial.PasswordEncrypted);
            }

            return await _userRepository.UpdateAsync(existing);
        }

        public async Task DeleteUser(long userUniqueId)
        {
            await _userRepository.DeleteAsync(userUniqueId);
        }
    }
}

