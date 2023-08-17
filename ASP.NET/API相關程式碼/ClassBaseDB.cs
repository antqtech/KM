using Microsoft.EntityFrameworkCore;
using System.Numerics;

namespace Execute_storedProcedure_DotnetCore.Models
{
    public class TestDbConext : DbContext
    {
        public TestDbConext(DbContextOptions<TestDbConext> options) : base(options)
        {

        }
        
        //範例
        public virtual DbSet<Player> Player { get; set; }
  
    }
}
