using System.Collections.Generic;
using System.Globalization;
using System.Runtime.Intrinsics.Arm;
using OfficeOpenXml;

class Program
{
    enum TableType{ 
        Spark = 0,
        VE = 1,
        AFR = 2
    }

    private static TableType outType;

    static void Main(string[] args)
    {
        Console.WriteLine("Enter the folder path containing the .csv files:");
        string folderPath = Console.ReadLine();

        if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
        {
            Console.WriteLine("Invalid folder path. Exiting.");
            return;
        }

        if(folderPath.Contains("Spark", StringComparison.CurrentCultureIgnoreCase)){ 
            outType = TableType.Spark;
        }
        else if(folderPath.Contains("VE", StringComparison.CurrentCultureIgnoreCase)){ 
            outType = TableType.VE;
        }
        else if(folderPath.Contains("AFR", StringComparison.CurrentCultureIgnoreCase)){ 
            outType = TableType.AFR;
        }

        List<int> mapBuckets = new();
        List<int> rpmBuckets = new();

        if(outType == TableType.VE){ 
            //Define the MAP and RPM buckets for VE data
            mapBuckets = new List<int> { 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 85, 95, 100 };
            rpmBuckets = new List<int> { 750, 1000, 1125, 1250, 1500, 1750, 2000, 2250, 2500, 2750, 3000, 3500, 4000, 4500, 5000, 5500, 6000, 6500, 7000, 7500, 8000 };
        }
        else if(outType == TableType.Spark){ 
            //Define the MAP and RPM buckets for Spark tables
            mapBuckets = new List<int> { 15, 20, 30, 40, 50, 60, 70, 80, 90, 95, 100 };
            rpmBuckets = new List<int> { 750, 1000, 1125, 1250, 1500, 1750, 2000, 2250, 2500, 2750, 3000, 3500, 4000, 4500, 5000, 5500, 6000, 6500, 7000, 7500, 8000 };
        }
        else if(outType == TableType.AFR){ 
        
        }
        
        // Buckets to store data
        var buckets = new Dictionary<(int Rpm, int Map), List<(double Front, double Rear)>>();

        // Initialize buckets
        foreach (var rpm in rpmBuckets)
        {
            foreach (var map in mapBuckets)
            {
                buckets[(rpm, map)] = new List<(double, double)>();
            }
        }

        // Process all .csv files in the folder
        var csvFiles = Directory.GetFiles(folderPath, "*.csv");
        foreach (var file in csvFiles)
        {
            Console.WriteLine($"Processing file: {file}");
            ProcessCsvFile(file, buckets, rpmBuckets, mapBuckets);
        }

        // Debugging: Log bucket contents before averaging
        Console.WriteLine("\nBucket Contents Before Averaging:");
        foreach (var bucket in buckets)
        {
            Console.WriteLine($"Bucket (RPM: {bucket.Key.Rpm}, MAP: {bucket.Key.Map}) contains {bucket.Value.Count} entries.");
            foreach (var entry in bucket.Value)
            {
                //Console.WriteLine($"  Front: {entry.Front}, Rear: {entry.Rear}");
            }
        }

        // Calculate averages for each bucket
        var averages = CalculateAverages(buckets);

        // Output results
        string outputFilePath = Path.Combine(folderPath, "ProcessedData.xlsx");
        WriteBucketsToExcel(outputFilePath, averages, mapBuckets, rpmBuckets);
        Console.WriteLine($"Processed data has been saved to: {outputFilePath}");
    }

    static void ProcessCsvFile(string filePath, Dictionary<(int Rpm, int Map), List<(double Front, double Rear)>> buckets, List<int> rpmBuckets, List<int> mapBuckets)
    {
        using (var reader = new StreamReader(filePath))
        {
            string headerLine = reader.ReadLine(); // Skip header
            if (headerLine == null) return;

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var columns = line.Split(',');

                if (columns.Length < 4) continue;

                if (!double.TryParse(columns[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double rpm) ||
                    !double.TryParse(columns[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double map) ||
                    !double.TryParse(columns[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double rear) ||
                    !double.TryParse(columns[3], NumberStyles.Float, CultureInfo.InvariantCulture, out double front))
                {
                    continue; // Skip invalid rows
                }

                // Determine the nearest MAP and RPM buckets
                int nearestMap = FindNearestBucket(map, mapBuckets);
                int nearestRpm = FindNearestBucket(rpm, rpmBuckets);

                // Add the data to the appropriate bucket
                buckets[(nearestRpm, nearestMap)].Add((front, rear));
            }
        }

        foreach (var mapVal in mapBuckets) {
            foreach (var rpmVal in rpmBuckets) {
                if (buckets[(rpmVal, mapVal)].Count < 10) {
                    buckets[(rpmVal, mapVal)].Clear();
                }
            }
        }
    }

    static int FindNearestBucket(double value, List<int> buckets)
    {
            for (int i = 0; i < buckets.Count - 1; i++) {
                double midpoint = (buckets[i] + buckets[i + 1]) / 2.0;
                if (value < midpoint) {
                    return buckets[i]; // Round down to the lower bucket
                }
            }

            return buckets.Last();
    }

    static Dictionary<(int Rpm, int Map), (double AvgFront, double AvgRear)> CalculateAverages(Dictionary<(int Rpm, int Map), List<(double front, double rear)>> buckets)
    {
        var averages = new Dictionary<(int Rpm, int Map), (double front, double rear)>();

        foreach (var bucket in buckets)
        {
            var (rpm, map) = bucket.Key;
            var list = bucket.Value;

            if(outType == TableType.VE){ 
                if (list.Count > 0){
                    double avgFront = list.Average(v => v.front);
                    double avgRear = list.Average(v => v.rear);
                    averages[(rpm, map)] = (avgFront, avgRear);
                }
                else{
                    averages[(rpm, map)] = (0, 0); // Default to 0 if no data
                }
            }
        
            if(outType == TableType.Spark){
                if (list.Count > 0){
                        double avgFront = list.Where(v => v.front != 0).Select(v => v.front).DefaultIfEmpty(0).Average();
                        double avgRear = list.Where(v => v.rear != 0).Select(v => v.rear).DefaultIfEmpty(0).Average();
                        averages[(rpm, map)] = (avgFront, avgRear);
                }
                else{
                        averages[(rpm, map)] = (0, 0); // Default to 0 if no data
                }
            }
        }

        return averages;
    }

    static void WriteBucketsToExcel(string filePath, Dictionary<(int Rpm, int Map), (double AvgFront, double AvgRear)> averages, List<int> mapBuckets, List<int> rpmBuckets)
    {
        if(File.Exists(filePath)){ 
                File.Delete(filePath);
            }

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            // Front worksheet
            var worksheetFront = package.Workbook.Worksheets.Add("Front");
            worksheetFront.Cells[1, 1].Value = "RPM / MAP";
            for (int i = 0; i < mapBuckets.Count; i++)
            {
                worksheetFront.Cells[1, i + 2].Value = mapBuckets[i];
            }
            for (int rowIndex = 0; rowIndex < rpmBuckets.Count; rowIndex++)
            {
                int rpm = rpmBuckets[rowIndex];
                worksheetFront.Cells[rowIndex + 2, 1].Value = rpm;
                for (int colIndex = 0; colIndex < mapBuckets.Count; colIndex++)
                {
                    int map = mapBuckets[colIndex];
                    if (averages.TryGetValue((rpm, map), out var values))
                    {
                        worksheetFront.Cells[rowIndex + 2, colIndex + 2].Value = values.AvgFront;
                    }
                    else
                    {
                        worksheetFront.Cells[rowIndex + 2, colIndex + 2].Value = 0;
                    }
                }
            }

            // Rear worksheet
            var worksheetRear = package.Workbook.Worksheets.Add("Rear");
            worksheetRear.Cells[1, 1].Value = "RPM / MAP";
            for (int i = 0; i < mapBuckets.Count; i++)
            {
                worksheetRear.Cells[1, i + 2].Value = mapBuckets[i];
            }
            for (int rowIndex = 0; rowIndex < rpmBuckets.Count; rowIndex++)
            {
                int rpm = rpmBuckets[rowIndex];
                worksheetRear.Cells[rowIndex + 2, 1].Value = rpm;
                for (int colIndex = 0; colIndex < mapBuckets.Count; colIndex++)
                {
                    int map = mapBuckets[colIndex];
                    if (averages.TryGetValue((rpm, map), out var values))
                    {
                        worksheetRear.Cells[rowIndex + 2, colIndex + 2].Value = values.AvgRear;
                    }
                    else
                    {
                        worksheetRear.Cells[rowIndex + 2, colIndex + 2].Value = 0;
                    }
                }
            }

            package.SaveAs(new FileInfo(filePath));
        }
    }
}
