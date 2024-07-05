using ILakshya.Dal;
using ILakshya.Model;
using ILakshya.WebApi.Jwt;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace ILakshya.WebApi.Controllers
{
    /*[Route("api/[controller]")]
    [ApiController]
    // bcrypt.net-next\4.0.3\,  'this package is used for hash pswd and varify password'
    public class WphAuthenticationController : ControllerBase
    {

      *//*  private readonly WebPocHubDbContext _dbContext;
        public WphAuthenticationController(WebPocHubDbContext context)
        {
            _dbContext = context;
        }
*//*

        private readonly IAuthenticationRepository _wphAuthentication;
        private readonly ITokenManager _tokenManager; // used in token manager for token

        public WphAuthenticationController(IAuthenticationRepository wphAuthentication, ITokenManager tokenManager) // token manager
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

    *//*    // In WphAuthenticationController.cs

        [HttpGet("CheckCredentials")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]

      *//*  public IActionResult Log(Student model)
        {
            var data = _dbContext.Students.Where(x => x.EnrollNo == model.EnrollNo).Select(x => x.Name).FirstOrDefault();
            if (data != null)
            {
                return Ok(data);
            }
            return Ok();

        }*//*
            public ActionResult<AuthResponses> GetDetails(Student user)
             {
                 try
                 {
                *//*if (user.RoleId == 1) // Admin
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
                }*//*
                // if (user.RoleId == 2) // Student
                {
                    //  var  data = WebPocHubDbContext.
                    //    string pass = "Ganesh@123";
                    // Assuming a common password for all students
                    Student obj = new Student();
                    obj.EnrollNo = user.EnrollNo;
                    if (obj.EnrollNo != null)
                    {
                        return Ok("Success");
                    }
                    *//*                         if (user.Password != "Ganesh56")
                                             {
                                                 return BadRequest("Incorrect Password! Please Check your Password");
                                             }*/

    /*                         var authUser = _wphAuthentication.CheckCredentials(user);
                             if (authUser == null)
                             {
                                 return NotFound();
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
                         }*//*
    else
    {
        return BadRequest("Invalid role");
    }
}
 }
 catch (InvalidOperationException ex)
 {
     return BadRequest(ex.Message);
 }
}



// ======================================= Update on 14 June


}
}*/

    [Route("api/[controller]")]
    [ApiController]
    public class WphAuthenticationController : ControllerBase
    {
        private readonly IAuthenticationRepository _wphAuthentication;
        private readonly ITokenManager _tokenManager; // used in token manager for token
        private readonly WebPocHubDbContext _dbContext;

        public WphAuthenticationController(IAuthenticationRepository wphAuthentication, ITokenManager tokenManager, WebPocHubDbContext dbContext)
        {
            _wphAuthentication = wphAuthentication;
            _tokenManager = tokenManager;
            _dbContext = dbContext;
        }

        [HttpPost("RegisterUser")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public ActionResult Create(User user)
        {
            try
            {
                if (user.RoleId == 1) // Admin
                {
                    // Admin registration
                    var passwordHash = BCrypt.Net.BCrypt.HashPassword(user.Password);
                    user.Password = passwordHash;
                    var result = _wphAuthentication.RegisterUser(user);
                    if (result > 0)
                    {
                        return Ok();
                    }
                }
                else if (user.RoleId == 2) // Student
                {
                    // Student registration with a single password for all enrollments
                    var passwordHash = BCrypt.Net.BCrypt.HashPassword("Ganesh56"); // Set a default password for students
                    var studentEnrollNos = _dbContext.Students.Select(s => s.EnrollNo).ToList();
                    if (studentEnrollNos.Count == 0)
                    {
                        return BadRequest("No student enrollments found.");
                    }
                    foreach (var enrollNo in studentEnrollNos)
                    {
                        var studentUser = new User
                        {
                            EnrollNo = enrollNo.ToString(),
                            Password = passwordHash,
                            RoleId = 2 // Student role
                        };
                        _wphAuthentication.RegisterUser(studentUser);
                    }
                    return Ok();
                }
                else
                {
                    return BadRequest("Invalid RoleId.");
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
                if (!BCrypt.Net.BCrypt.Verify(user.Password, authUser.Password))
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
        }
    }
}