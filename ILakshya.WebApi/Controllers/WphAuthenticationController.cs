using ILakshya.Dal;
using ILakshya.Model;
using ILakshya.WebApi.Jwt;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ILakshya.WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    // bcrypt.net-next\4.0.3\,  'this package is used for hash pswd and varify password'
    public class WphAuthenticationController : ControllerBase
         { 
            private readonly IAuthenticationRepository _wphAuthentication;
            private readonly ITokenManager _tokenManager; // used in token manager for token
            public WphAuthenticationController(IAuthenticationRepository wphAuthentication,ITokenManager tokenManager) // token manager
            {
                _wphAuthentication = wphAuthentication;
                _tokenManager = tokenManager; // used token manager
            }
            [HttpPost("RegisterUser")]
            [ProducesResponseType(StatusCodes.Status200OK)]
            [ProducesResponseType(StatusCodes.Status400BadRequest)]
            public ActionResult Create(User user)
            {
                var passwordHash = BCrypt.Net.BCrypt.HashPassword(user.Password);
                user.Password = passwordHash;

            try
            {
                var result = _wphAuthentication.RegisterUser(user);


                if (result > 0)
                {
                    return Ok();
                }
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
            return BadRequest();
            }

            [HttpPost("CheckCredentials")]
            [ProducesResponseType(StatusCodes.Status200OK)]
            [ProducesResponseType(StatusCodes.Status404NotFound)]
            [ProducesResponseType(StatusCodes.Status400BadRequest)]
            public ActionResult<AuthResponses> GetDetails(User user)
            {
            try
            {
                var authUser = _wphAuthentication.CheckCredentials(user);
                if (authUser == null)
                {
                    return NotFound();
                }
                if (authUser != null && !BCrypt.Net.BCrypt.Verify(user.Password, authUser.Password))
                {
                    return BadRequest("Incorrect Password! Please Check your Password");
                }
                var roleName = _wphAuthentication.GetUserRole(authUser.RoleId);

                var authResponse = new AuthResponses()
                {
                    IsAuthenticated = true,
                    Role = roleName,
                    Token = _tokenManager.GenerateToken(authUser, roleName),
                    EnrollNo = authUser.EnrollNo
                };  
                return Ok(authResponse);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
                //     {
                //    IsAuthenticated = true,
                //    Role = roleName,
                //    Token = _tokenManager.GenerateToken(authUser, roleName)
                //}; // token manager used GenerateToken(authUser)
                //return Ok(authResponse);
            }
        }
    }