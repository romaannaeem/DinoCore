using DinoCore.Models;
using Microsoft.EntityFrameworkCore;

namespace DinoCore.Data
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
        {
        }

        public DbSet<Dinosaur> Dinosaurs => Set<Dinosaur>();
    }
}
