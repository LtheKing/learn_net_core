using Application.Students;
using Domain;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Persistence;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ContosoUniversity.Controllers.v1
{
    public class StudentController : BaseAPIController
    {
        private readonly DataContext _context;

        public StudentController(DataContext context)
        {
            _context = context;
        }
        [HttpGet]
        public async Task<ActionResult<List<Student>>> GetStudents()
        {
            //var result = await Mediator.Send(new List.Query());
            var result = _context.Students.ToListAsync();
            return Ok(result);
        }
    }
}
