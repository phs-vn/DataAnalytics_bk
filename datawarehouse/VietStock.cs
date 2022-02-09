using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Dynamic;
using System.Data;

namespace DynamicCS
   {
    class VietStock
    {
        public DataTable GetVietStockData(int group, int type, DateTime from, DateTime to)
        {
            using (DataWebServiceClient client = new DataWebServiceClient())
            {
                return client.GetData(group,type,from,to); 
            }
        }

        DataTable p = new GetVietStockData(
        4, 6, DateTime(2022,1,25), DateTime(2022,1,25)
        );

        static void Main(string[] args)
        {Console.WriteLine(p);}
    }
    }

    


