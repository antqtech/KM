using System.ComponentModel.DataAnnotations;
namespace Execute_storedProcedure_DotnetCore.Models
{
    //範例
    public class Player
    {
        [Key]
        public string PID { get; set; }
        public string PName_Zh { get; set; }
        public string PName_En { get; set; }
    }

   
}
