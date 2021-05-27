using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanMakeTree.DataBase
{
    public class ChemicalContext: DbContext
    {
        public DbSet<Molecule> Molecule { get; set; }

        protected override void OnConfiguring(
            DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseNpgsql("User ID=postgres;Password=postgres;Server=192.168.0.143;Port=5432;Database=chemicalinfo3.5");
            optionsBuilder.UseLazyLoadingProxies();
            base.OnConfiguring(optionsBuilder);
        }
    }
}
