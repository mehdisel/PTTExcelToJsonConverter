using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace PTTExcelToJsonConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            /*Bu kısım dinamik olarak çekilebilir.*/
            var filePath = Directory.GetFiles(Environment.CurrentDirectory).ToList().FirstOrDefault(x=>x.EndsWith(".xlsx"));
            if (string.IsNullOrWhiteSpace(filePath))
                Console.WriteLine(".xlsx uzantılı pk listesi bulunamadı.");
            /*kolon isminden dinamik olarak yakalanabilir.*/
            var cityColumnIndex = 1;
            var districtColumnIndex = 2;
            var streetColumnIndex = 4;
            var pkColumnIndex = 5;

            using FileStream fileStream = File.OpenRead(filePath);
            
            var workBook=new XLWorkbook(fileStream);
            var workSheet = workBook.Worksheet(1);
            int lastRowNumber = workSheet.LastRowUsed().RowNumber();
            var rows = workSheet.Rows(2, lastRowNumber);

            var cityList = new List<Models.City>();
            int districtIdIndex = 1;
            int streetIdIndex = 1;
            foreach (var row in rows)
            {
                var cityCell = row.Cell(cityColumnIndex);
                if (!cityCell.TryGetValue<string>(out string city))
                    throw new Exception(cityCell.Address.ToString());

                var cityIdIndex = cityList.Count() > 0 ? cityList.Max(x => x.Id) +1 : 1;
                var beforeAddedCity = cityList.FirstOrDefault(x => x.Name == city.Trim()) ?? new Models.City { 
                    Id = cityIdIndex,
                    Name=city.Trim()
                };

                var districtCell = row.Cell(districtColumnIndex);
                if (!districtCell.TryGetValue<string>(out string district))
                    throw new Exception(districtCell.Address.ToString());
                var beforeAddedDistrict = beforeAddedCity.Districts.FirstOrDefault(x => x.Name == district.Trim()) ?? new Models.District
                {
                    Id = districtIdIndex,
                    Name = district.Trim(),
                    CityId = beforeAddedCity.Id,
                    City = beforeAddedCity
                };


                if (!beforeAddedCity.Districts.Any(x => x.Id == beforeAddedDistrict.Id))
                { 
                    beforeAddedCity.Districts.Add(beforeAddedDistrict);
                    districtIdIndex++;
                }

                var streetCell = row.Cell(streetColumnIndex);
                if (!streetCell.TryGetValue<string>(out string street))
                    throw new Exception(streetCell.Address.ToString());
                var pkCell = row.Cell(pkColumnIndex);
                if (!pkCell.TryGetValue<string>(out string pk))
                    throw new Exception(pkCell.Address.ToString());

                beforeAddedDistrict.Streets.Add(new Models.Street
                {
                    City = beforeAddedCity,
                    District = beforeAddedDistrict,
                    CityId = beforeAddedCity.Id,
                    DistrictId = beforeAddedDistrict.Id,
                    Id = streetIdIndex,
                    PK = pk,
                    Name = street.Trim()
                });
                streetIdIndex++;

                if (!cityList.Any(x=>x.Id==beforeAddedCity.Id))
                cityList.Add(beforeAddedCity);
            }
            var citiesList = cityList.Select(x => new
            {
                Id = x.Id,
                Name = x.Name
            }).ToList();
            var districtsList = cityList.SelectMany(x=> x.Districts).Select(x=> new
            {
                Id=x.Id,
                Name = x.Name,
                CityId=x.CityId
            }).ToList();

            var streetsList = cityList.SelectMany(x => x.Districts.SelectMany(y => y.Streets)).Select(x=>
            new {
                Id=x.Id,
                Name=x.Name,
                DistrictId= x.DistrictId,
                CityId= x.CityId,
                PK=x.PK
            }).ToList();

            var convertedPath = Path.Combine(Environment.CurrentDirectory, "Converted");
            var convertedPathCreated = Directory.Exists(convertedPath);
            if (!convertedPathCreated)
                Directory.CreateDirectory(convertedPath);
                    
            
            
            File.WriteAllText(Path.Combine(convertedPath, "convertedAll.json") ,Newtonsoft.Json.JsonConvert.SerializeObject(cityList,Newtonsoft.Json.Formatting.None,
                        new JsonSerializerSettings()
                        {
                            ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                        }));
            File.WriteAllText(Path.Combine(convertedPath,"cities.json") ,Newtonsoft.Json.JsonConvert.SerializeObject(citiesList));
            File.WriteAllText(Path.Combine(convertedPath,"districts.json") ,Newtonsoft.Json.JsonConvert.SerializeObject(districtsList));
            File.WriteAllText(Path.Combine(convertedPath,"streets.json") ,Newtonsoft.Json.JsonConvert.SerializeObject(streetsList));

            Process.Start("explorer.exe", convertedPath);
            Console.WriteLine("Dönüşüm tamamlandı");
            Console.Read();
        }
    }
}
