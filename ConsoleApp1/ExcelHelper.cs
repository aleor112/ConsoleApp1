using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestProject1;

namespace ConsoleApp1
{
    public class ExcelHelper
    {

        public static void GenerateExcel(List<TestProject1.Record> records, string suffix)
        {
            // Instantiate a new Workbook
            Workbook book = new Workbook();
            // Obtaining the reference of the worksheet
            Worksheet sheet = book.Worksheets[0];
            InsertHeaderRow(sheet);
             sheet.Cells.ImportCustomObjects((System.Collections.ICollection)records,
                 new string[] {nameof(Record.TicketId),nameof(Record.TicketType),nameof(Record.Application),nameof(Record.Service),nameof(Record.ServiceOffering),
                     nameof(Record.Priority),nameof(Record.Status),nameof(Record.TicketOpenedAt),
                     nameof(Record.TicketClosedAt),nameof(Record.TicketForwardedAt), nameof(Record.LastAdessoSupporter), nameof(Record.AllConnectedAdessoSupporters),
                     nameof(Record.AssignmentTime),nameof(Record.ReactionTime), nameof(Record.CurrentAssignmentGroup), nameof(Record.FirstAdessoAssignmentGroup), nameof(Record.LastAdessoAssignmentGroup), 
                     nameof(Record.RelatedWithCloudMove), nameof(Record.RelatedWithRepairMonitor),
                     nameof(Record.RelatedWithLTS), nameof(Record.RelatedWithInternalApplications), nameof(Record.RelatedWithCCPT), 
                     nameof(Record.RelatedWithCapacityTool),nameof(Record.RelatedWithBluenet), nameof(Record.RelatedAssignmentGroupsCount)
                },
                false, // isPropertyNameShown
                1, // firstRow
                0, // firstColumn
                records.Count, // Number of objects to be exported
                true, // insertRows
                null, // dateFormatString
                false); // convertStringToNumber

            // Save the Excel file
            book.Save($"ticketReport_time_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}_fromRow_{suffix}.xlsx");
        }


        static void InsertHeaderRow(Worksheet worksheet)
        {
            // Define header row data
            string[] headers = { "Ticket Id","Ticket type","Application","Service","Service offering","Prio","Status","Ticket opened at","Ticket closed at","Ticket forwarded at", "Last adesso supporter", "All connected adesso supporters", "adesso assignment time",
                    "adesso first response time", "current assignment group", "first adesso assignment group", "last adesso assignment group", "related with SP 2013 Cloud Move", "related with Repair Monitor",
                "related with LTS", "related with internal applications", "related with ccpt", "related with capacity tool","related with bluenet", "count of related Assigment Groups" };

            // Insert header row at A1
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[0, i].PutValue(headers[i]);
            }
        }
    }
}
