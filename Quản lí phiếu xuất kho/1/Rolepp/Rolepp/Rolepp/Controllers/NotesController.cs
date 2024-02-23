using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using Rolepp.Data;
using Rolepp.Models;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Rolepp.Controllers
{
    public class NotesController : Controller
    {
        private readonly ApplicationDbContext _context;

        public NotesController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Notes

        public async Task<IActionResult> Index(string sortOrder, int? pageNumber, string searchString, DateTime? fromDate, DateTime? toDate)
        {

            IQueryable<Note> notes = _context.Notes
                .Include(n => n.NoteProducts)
                .ThenInclude(np => np.Product);

            switch (sortOrder)
            {
                case "newest":
                    notes = notes.OrderByDescending(s => s.CreatedDate);
                    break;
                case "oldest":
                    notes = notes.OrderBy(s => s.CreatedDate);
                    break;
            }
            if (!String.IsNullOrEmpty(searchString))
            {
                notes = notes.Where(s => s.NoteCode.Contains(searchString));
            }
            if (fromDate.HasValue && toDate.HasValue)
            {
                notes = notes.Where(s => s.CreatedDate >= fromDate && s.CreatedDate <= toDate);
            }
            if (fromDate.HasValue && toDate.HasValue)
            {
                notes = notes.Where(s => s.CreatedDate >= fromDate && s.CreatedDate <= toDate);
                TempData["FromDate"] = fromDate.Value.ToString("yyyy-MM-dd");
                TempData["ToDate"] = toDate.Value.ToString("yyyy-MM-dd");
            }

            int pageSize = 10;
            return View(await PaginatedList<Note>.CreateAsync(notes.AsNoTracking(), pageNumber ?? 1, pageSize));
        }


        // GET: Notes/Create
        public IActionResult Create()
        {
            // Retrieve products for dropdown
            var products = _context.Products.ToList();

            // Pass products to the ViewBag
            ViewBag.Products = products;

            return View();
        }

        // POST: Notes/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(NoteViewModel model)
        {

            var duplicateProducts = model.Products
             .GroupBy(p => p.ProductID)
             .Where(g => g.Count() > 1)
             .Select(g => g.Key)
             .ToList();


            // Kiểm tra sự tồn tại của sản phẩm trong cơ sở dữ liệu
            foreach (var productViewModel in model.Products)
            {
                var productInDb = _context.Products.FirstOrDefault(p => p.ProductID == productViewModel.ProductID);
                if (productInDb == null)
                {
                    ModelState.AddModelError("", "Invalid product selection.");
                    return View(model);
                }
            }
            if (ModelState.IsValid)
            {
                // Map NoteViewModel to Note entity
                var note = new Note
                {
                    NoteCode = model.NoteCode,
                    CreateName = model.CreateName,
                    Customer = model.Customer,
                    AddressCustomer = model.AddressCustomer,
                    Reason = model.Reason,
                    Status = 1,
                    CreatedDate = DateTime.Now,
                    Total = 0
                };

                _context.Notes.Add(note);
                _context.SaveChanges();

                foreach (var productViewModel in model.Products)
                {
                    var noteProduct = new NoteProduct
                    {
                        NoteId = note.NoteId, // Set NoteId from the saved note
                        ProductID = productViewModel.ProductID,
                        StockOut = productViewModel.StockOut
                        // Optionally, you can set other properties of NoteProduct here
                    };

                    _context.NoteProducts.Add(noteProduct);
                }
                foreach (var product in model.Products)
                {
                    var productInDb = _context.Products.FirstOrDefault(p => p.ProductID == product.ProductID);
                    if (productInDb != null)
                    {
                        productInDb.Quantity -= product.StockOut; // Giảm số lượng sản phẩm bằng số lượng nhập
                    }
                }

                _context.SaveChanges(); // Save changes to save NoteProducts

                return RedirectToAction(nameof(Index));
            }

            // If ModelState is not valid, return to the create view with errors

            return View(model);
        }


        // GET: Notes/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var note = await _context.Notes.FindAsync(id);
            if (note == null)
            {
                return NotFound();
            }
            return View(note);
        }

        // POST: Notes/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("NoteId,NoteCode,CreateName,Customer,AddressCustomer,Reason,Status")] Note note)
        {
            if (id != note.NoteId)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(note);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!NoteExists(note.NoteId))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(note);
        }

        // GET: Details
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var note = await _context.Notes
                .Include(n => n.NoteProducts)
                .ThenInclude(np => np.Product) // Include Product information
                .FirstOrDefaultAsync(m => m.NoteId == id);

            if (note == null)
            {
                return NotFound();
            }

            return View(note);
        }



        // GET: Notes/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var note = await _context.Notes
                .FirstOrDefaultAsync(m => m.NoteId == id);
            if (note == null)
            {
                return NotFound();
            }

            return View(note);
        }

        // POST: Notes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var note = await _context.Notes.FindAsync(id);
            _context.Notes.Remove(note);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        public async Task<IActionResult> UpdateStatusAjax(int id, int newStatus)
        {
            var note = await _context.Notes.FindAsync(id);

            if (note == null)
            {
                return NotFound();
            }

            note.UpdateStatus(newStatus);

            _context.Entry(note).State = EntityState.Modified;

            await _context.SaveChangesAsync();

            return Ok();
        }

        private bool NoteExists(int id)
        {
            return _context.Notes.Any(e => e.NoteId == id);
        }
        public IActionResult GetNewNoteCount()
        {
            int newNoteCount = _context.Notes.Count(n => n.Status == 2);
            return Json(newNoteCount);
        }

        public IActionResult CheckNoteStatus()
        {
            bool hasNoteStatus34 = _context.Notes.Any(n => n.Status == 3 || n.Status == 4);
            return Json(hasNoteStatus34);
        }

        public IActionResult DownloadNoteDetails(int id)
        {
            var note = _context.Notes.Include(n => n.NoteProducts).ThenInclude(np => np.Product).FirstOrDefault(n => n.NoteId == id);
            if (note == null)
            {
                return NotFound();
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var worksheet = package.Workbook.Worksheets.Add("Note Details");

                // Add Note details
                worksheet.Cells[1, 3].Value = "Note Delivery Goods Information";
                worksheet.Cells[2, 1].Value = "Note Code";
                worksheet.Cells[2, 2].Value = note.NoteCode;

                worksheet.Cells[3, 1].Value = "Created's Name";
                worksheet.Cells[3, 2].Value = note.CreateName;

                worksheet.Cells[4, 1].Value = "Customer";
                worksheet.Cells[4, 2].Value = note.Customer;

                worksheet.Cells[5, 1].Value = "Customer's Address";
                worksheet.Cells[5, 2].Value = note.AddressCustomer;

                worksheet.Cells[6, 1].Value = "Reason";
                worksheet.Cells[6, 2].Value = note.Reason;

                worksheet.Cells[7, 1].Value = "Date Created";
                worksheet.Cells[7, 2].Value = note.CreatedDate.ToString();
                worksheet.Cells[9, 3].Value = "Product to Export";
                // Add header for Products
                worksheet.Cells[10, 1].Value = "Product Name";
                worksheet.Cells[10, 2].Value = "Product Code";
                worksheet.Cells[10, 3].Value = "StockOut";
                worksheet.Cells[10, 4].Value = "Price";
                worksheet.Cells[10, 5].Value = "Total";

                // Add data for Products
                int row = 11;
                foreach (var product in note.NoteProducts)
                {
                    worksheet.Cells[row, 1].Value = product.Product.ProductName;
                    worksheet.Cells[row, 2].Value = product.Product.ProductCode;
                    worksheet.Cells[row, 3].Value = product.StockOut;
                    worksheet.Cells[row, 4].Value = product.Product.Price;
                    worksheet.Cells[row, 5].Value = product.StockOut * product.Product.Price;
                    row++;
                }

                // Add total of Note
                worksheet.Cells[row, 4].Value = "Total of Note:";
                worksheet.Cells[row, 5].Value = note.NoteProducts.Sum(p => p.StockOut * p.Product.Price);

                // Save the Excel package to a MemoryStream
                var stream = new MemoryStream();
                package.SaveAs(stream);

                // Return the Excel file as a FileContentResult
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Note Code: " + note.NoteCode + ".xlsx");
            }
        }

        public IActionResult DownloadSearchResults()
        {
            DateTime fromDate = DateTime.Parse(TempData["FromDate"].ToString());
            DateTime toDate = DateTime.Parse(TempData["ToDate"].ToString());

            IQueryable<Note> notes = _context.Notes.Include(n => n.NoteProducts).ThenInclude(np => np.Product);

            notes = notes.Where(s => s.CreatedDate.Date >= fromDate.Date && s.CreatedDate.Date <= toDate.Date);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var worksheet = package.Workbook.Worksheets.Add("Search Results");

                // Add header
                worksheet.Cells[1, 1].Value = "Note Code";
                worksheet.Cells[1, 2].Value = "Created's Name";
                worksheet.Cells[1, 3].Value = "Customer";
                worksheet.Cells[1, 4].Value = "Customer's Address";
                worksheet.Cells[1, 5].Value = "Reason";
                worksheet.Cells[1, 6].Value = "Date Created";
                worksheet.Cells[1, 7].Value = "Products and StockOut"; // New column
                worksheet.Cells[1, 8].Value = "Total";

                // Add data
                int row = 2;
                foreach (var note in notes)
                {
                    worksheet.Cells[row, 1].Value = note.NoteCode;
                    worksheet.Cells[row, 2].Value = note.CreateName;
                    worksheet.Cells[row, 3].Value = note.Customer;
                    worksheet.Cells[row, 4].Value = note.AddressCustomer;
                    worksheet.Cells[row, 5].Value = note.Reason;
                    worksheet.Cells[row, 6].Value = note.CreatedDate.ToString();

                    // Create a string containing all products and their StockOut, separated by newlines
                    var productsAndStockOut = string.Join("\n", note.NoteProducts.Select(np => np.Product.ProductName + ": " + np.StockOut));
                    worksheet.Cells[row, 7].Value = productsAndStockOut;

                    worksheet.Cells[row, 8].Value = note.NoteProducts.Sum(p => p.StockOut * p.Product.Price);
                    row++;
                }

                // Save the Excel package to a MemoryStream
                var stream = new MemoryStream();
                package.SaveAs(stream);

                // Return the Excel file as a FileContentResult
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SearchResults.xlsx");
            }
        }

    }
}
