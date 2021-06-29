using MediatR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Domain;
using Persistence;
using System.Threading;
using Microsoft.EntityFrameworkCore;

namespace Application.Students
{
    public class List
    {
        public class Query : IRequest<List<Student>> { }

        public class Handler : IRequestHandler<Query, List<Student>>
        {
            private readonly DataContext _context;

            public Handler(DataContext context)
            {
                _context = context;
            }

            public async Task<List<Student>> Handle(Query request, CancellationToken cancellationToken)
            {
                return await _context.Students.ToListAsync();
            }
        }
    }
}
