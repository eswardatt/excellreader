using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1
{
    class LinqJoins
    {
    }
    public class Demo
    {
        List<Claim> claims = new List<Claim>()
            {
               new  Claim{  ClaimId=1,IsApproved=true,CID="1_Ind"},
               new  Claim{  ClaimId=2,IsApproved=true,CID="2_USA"},
               new  Claim{  ClaimId=3,IsApproved=false,CID="3_China"},
               new  Claim{  ClaimId=4,IsApproved=false,CID="4_Japan"},
               new  Claim{  ClaimId=5,IsApproved=true,CID="4_UK"},
            };


        List<Country> coutries = new List<Country>()
            {
               new Country{CID="1",CName="India"},
               new Country{CID="2",CName="US of America"},
               new Country{CID="3",CName="China"},
               new Country{CID="4",CName="Japan"},
               new Country{CID="5",CName="UK"},
                           };

        public void Get()
        {
            //foreach (var item in claims)
            //{
            //  item.CID = item.CID.Substring(0, item.CID.IndexOf('_'));
            //}
            //List<Claim> claims_1 = claims;
            var res = claims.Join(coutries, a => a.CID.Substring(0, a.CID.IndexOf('_')), b => b.CID, (a, b) => new { a.ClaimId, b.CName, a.IsApproved }).ToList().AsQueryable().Where(a => a.IsApproved == true);
            foreach (var item in res)
            {
                Console.WriteLine("ClaimId : {0} IsApproved : {1} Country {2}", item.ClaimId, item.IsApproved, item.CName);
            }
        }
    }
    public class Claim
    {
        public int ClaimId { get; set; }
        public bool IsApproved { get; set; }
        public string CID { get; set; }
    }

    public class Country
    {
        public string CID { get; set; }
        public string CName { get; set; }
    }
}
