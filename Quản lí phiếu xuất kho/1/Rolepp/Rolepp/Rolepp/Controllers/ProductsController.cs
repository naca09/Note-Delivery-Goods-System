using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Build.Evaluation;
using Microsoft.CodeAnalysis;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using Rolepp.Data;
using Rolepp.Models;
using System.Linq;
using System.Reflection.Metadata;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Rolepp.Controllers
{
    public class ProductsController : Controller
    {
        private readonly ApplicationDbContext _context;

        public ProductsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Product
        public async Task<IActionResult> Index(string searchString, int? pageNumber, bool lowQuantity = false)
        {
            // Retrieve data from the ProductsTest table
            var products = _context.Products.AsQueryable();

            if (!string.IsNullOrEmpty(searchString))
            {
                products = products.Where(s => s.ProductName.Contains(searchString));
            }

            if (lowQuantity)
            {
                products = products.Where(p => p.Quantity < 10);
            }

            products = products.Take(1000); // Limit to top 1000 rows

            // Include Warehouse information after all filters are applied
            var productsWithWarehouse = products
                .Include(p => p.Warehouse)
                .AsNoTracking();

            int pageSize = 10;
            return View(await PaginatedList<Product>.CreateAsync(productsWithWarehouse, pageNumber ?? 1, pageSize));
        }





        // GET: Product/Create
        public IActionResult Create()
        {
            // Get the list of warehouses to populate the dropdown
            ViewBag.Warehouses = new SelectList(_context.Warehouses, "WarehouseId", "WarehouseName");
            return View();
        }

        // POST: Product/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("ProductId,ProductName,Price,Quantity,ProductCode,WarehouseId")] Product product)
        {
            if (product.ProductID != null && product.ProductName != null && product.Price != null && product.Quantity != null && product.ProductCode != null && product.WarehouseId != null && product.Warehouse == null)
            {
                _context.Add(product);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }

            ViewBag.Warehouses = new SelectList(_context.Warehouses, "WarehouseId", "WarehouseName", product.WarehouseId);
            return View(product);

        }

        // GET: Product/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var product = await _context.Products.FindAsync(id);
            if (product == null)
            {
                return NotFound();
            }

            // Get the list of warehouses to populate the dropdown
            ViewBag.Warehouses = new SelectList(_context.Warehouses, "WarehouseId", "WarehouseName", product.WarehouseId);
            return View(product);
        }

        // POST: Product/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("ProductID,ProductName,Price,Quantity,ProductCode,WarehouseId")] Product product)
        {
            if (id != product.ProductID)
            {
                return NotFound();
            }

            if (product.ProductID != null && product.ProductName != null && product.Price != null && product.Quantity != null && product.ProductCode != null && product.WarehouseId != null && product.Warehouse == null)
            {
                try
                {
                    _context.Update(product);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!ProductExists(product.ProductID))
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

            // If validation fails, re-populate the dropdown with the list of warehouses
            ViewBag.Warehouses = new SelectList(_context.Warehouses, "WarehouseId", "WarehouseName", product.WarehouseId);
            return View(product);
        }

        // GET: Product/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var product = await _context.Products
                .FirstOrDefaultAsync(m => m.ProductID == id);
            if (product == null)
            {
                return NotFound();
            }

            return View(product);
        }
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var product = await _context.Products
                .Include(p => p.Warehouse) // Include Warehouse information
                .FirstOrDefaultAsync(m => m.ProductID == id);

            if (product == null)
            {
                return NotFound();
            }

            return View(product);
        }
        // POST: Product/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var product = await _context.Products.FindAsync(id);
            _context.Products.Remove(product);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool ProductExists(int id)
        {
            return _context.Products.Any(e => e.ProductID == id);
        }
        // Trong Controller của bạn


        public IActionResult Search(string searchText)
        {
            // Logic để tìm kiếm sản phẩm dựa trên searchText
            var products = _context.Products.Where(p => p.ProductName.Contains(searchText)).ToList();
            return View("Index", products);
        }

        [HttpGet]
        public JsonResult GetProductDetails(int id)
        {
            var product = _context.Products.FirstOrDefault(p => p.ProductID == id);
            if (product != null)
            {
                return Json(new { productCode = product.ProductCode, quantity = product.Quantity, price = product.Price });
            }
            return Json(null);
        }



        [HttpGet]
        public IActionResult ProductCount()
        {
            var lowQuantityProductCount = _context.Products.Count(p => p.Quantity < 10);
            return Json(lowQuantityProductCount);
        }


        public async Task<IActionResult> Download()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Products");
                worksheet.Cells["A1"].Value = "Product's Name";
                worksheet.Cells["B1"].Value = "Price";
                worksheet.Cells["C1"].Value = "Quantity";
                worksheet.Cells["D1"].Value = "Product Code";
                worksheet.Cells["E1"].Value = "Warehouse";

                var products = await _context.Products.Include(p => p.Warehouse).ToListAsync(); // Include Warehouse
                for (int i = 0; i < products.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = products[i].ProductName;
                    worksheet.Cells[i + 2, 2].Value = products[i].Price;
                    worksheet.Cells[i + 2, 3].Value = products[i].Quantity;
                    worksheet.Cells[i + 2, 4].Value = products[i].ProductCode;
                    worksheet.Cells[i + 2, 5].Value = products[i].Warehouse.WarehouseName; // Use WarehouseName
                }

                return File(package.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Products.xlsx");
            }
        }







    }
}