using ExcelCore.Models;
using Microsoft.EntityFrameworkCore;
namespace ExcelCore
{
    public class DBCtx : DbContext
    {
        public DBCtx()
        {
        }
        public DBCtx(DbContextOptions<DBCtx> option) : base(option) { }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Info>().HasNoKey();
        }

        public DbSet<Info> Infos { get; set; }
    }
}
