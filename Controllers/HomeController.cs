using Microsoft.AspNetCore.Mvc;
using PPTXChart.Models;
using System.Diagnostics;
using Syncfusion.Presentation;
using Syncfusion.OfficeChart;

namespace PPTXChart.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult BarChart()
        {
            IPresentation presentation = Presentation.Create();
            CreateBarChart(presentation);
            MemoryStream stream = new MemoryStream();
            presentation.Save(stream);
            return File(stream, "application/presentation", "Barchart.pptx");

        }

        private void CreateBarChart(IPresentation presentation)
        {
            ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
            IPresentationChart barchart = slide.Shapes.AddChart(140, 30, 680, 480);
            barchart.ChartType = OfficeChartType.Bar_Clustered;
            AddBarChartData(barchart);
            barchart.ChartTitle = "Purchase Details";
            barchart.ChartTitleArea.FontName = "Calibri";
            barchart.ChartTitleArea.Size = 14;
            IOfficeChartSerie serie1 = barchart.Series.Add("Sum of Future Expenses");
            serie1.Values = barchart.ChartData[2, 2, 6, 2];
            IOfficeChartSerie serie2 = barchart.Series.Add("Sum of Purchases");
            serie2.Values = barchart.ChartData[2, 3, 6, 3];

            barchart.HasDataTable= true;
            barchart.DataTable.HasBorders= true; 
            barchart.DataTable.HasHorzBorder = true;
            barchart.DataTable.HasVertBorder= true;
            barchart.DataTable.ShowSeriesKeys= true;
            barchart.HasLegend = false;

            barchart.PlotArea.Border.LinePattern = OfficeChartLinePattern.Solid;
            barchart.ChartArea.Border.LinePattern = OfficeChartLinePattern.Solid;
            barchart.PrimaryCategoryAxis.CategoryLabels = barchart.ChartData[2, 1, 6, 1];
            barchart.PrimaryCategoryAxis.Font.Size = 12;
            barchart.PrimaryCategoryAxis.MajorTickMark = OfficeTickMark.TickMark_None;
            barchart.PrimaryValueAxis.MajorTickMark = OfficeTickMark.TickMark_None;
        }

        private void AddBarChartData(IPresentationChart chart)
        {
            chart.ChartData.SetValue(1, 2, "Sum of Future Expenses");
            chart.ChartData.SetValue(1, 3, "Sum of Purchases");
            chart.ChartData.SetValue(2, 1, "Nancy Davalio");
            chart.ChartData.SetValue(2, 2, 1300);
            chart.ChartData.SetValue(2, 3, 600);
            chart.ChartData.SetValue(3, 1, "Andrew Fuller");
            chart.ChartData.SetValue(3, 2, 680);
            chart.ChartData.SetValue(3, 3, 1000);
            chart.ChartData.SetValue(4, 1, "Janet Leverling");
            chart.ChartData.SetValue(4, 2, 1280);
            chart.ChartData.SetValue(4, 3, 800);
            chart.ChartData.SetValue(5, 1, "Margaret Peacock");
            chart.ChartData.SetValue(5, 2, 2000);
            chart.ChartData.SetValue(5, 3, 400);
            chart.ChartData.SetValue(6, 1, "Steven Buchanan");
            chart.ChartData.SetValue(6, 2, 2660);
            chart.ChartData.SetValue(6, 3, 731);
        }

        public IActionResult PieChart()
        {
            IPresentation presentation = Presentation.Create();
            CreatePieChart(presentation);
            MemoryStream stream = new MemoryStream();
            presentation.Save(stream);
            return File(stream, "application/presentation", "Piechart.pptx");

        }

        private void CreatePieChart(IPresentation presentation)
        {
            ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
            IPresentationChart pieChart = slide.Shapes.AddChart(150, 80, 550, 400);
            pieChart.ChartType = OfficeChartType.Pie;
            AddPieChartData(pieChart);
            pieChart.DataRange = pieChart.ChartData[2, 1, 7, 2];
            pieChart.IsSeriesInRows = false;
            pieChart.ChartTitle = "Car Sales";
            pieChart.HasLegend = true;
            pieChart.Legend.Position = OfficeLegendPosition.Bottom;
            IOfficeChartSerie serie = pieChart.Series[0];
            serie.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
            serie.SerieFormat.LineProperties.LinePattern = OfficeChartLinePattern.Solid;
            serie.SerieFormat.LineProperties.LineColor = Syncfusion.Drawing.Color.White;

        }

        private void AddPieChartData(IPresentationChart chart)
        {
            chart.ChartData.SetValue(2, 1, "Car Types");
            chart.ChartData.SetValue(2, 2, "Units");
            chart.ChartData.SetValue(3, 1, "Mid size Cars");
            chart.ChartData.SetValue(3, 2, 85000);
            chart.ChartData.SetValue(4, 1, "Compact Car");
            chart.ChartData.SetValue(4, 2, 84000);
            chart.ChartData.SetValue(5, 1, "Compact SUV");
            chart.ChartData.SetValue(5, 2, 205000);
            chart.ChartData.SetValue(6, 1, "Full size Truck");
            chart.ChartData.SetValue(6, 2, 190000);
            chart.ChartData.SetValue(7, 1, "Mid size SUV");
            chart.ChartData.SetValue(7, 2, 225000);
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}