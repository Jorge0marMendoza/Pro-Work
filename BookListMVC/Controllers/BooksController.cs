using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BookListMVC.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace BookListMVC.Controllers
{
    public class BooksController : Controller
    {
        //create a local variable to assign the DB object to
        private readonly ApplicationDbContext _db;

        //retieve object that exist for us
        //create constructor with the db object as a argument 
        public BooksController(ApplicationDbContext db)
        {
            _db = db;
        }

        public IActionResult Index()
        {
            return View(_db.Books.ToList());
        }

        //GET-retrives or loads the page
        //Get BOOK/CREATE
        public IActionResult Create()
        {
            return View();
        }

        //Post-fills in the deatils and pass the values to create new
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Book book)
        {
            if (ModelState.IsValid) //In the model.book we defined the book properties that were required
            {
                _db.Add(book); //adds book inside a queue 
                await _db.SaveChangesAsync(); //this actually save the changes to the database
                return RedirectToAction("Index"); //take the user back the ther index book list to show the new list of books 
            }
            return View(book);
        }

        //GET ACTION METHOD FOR EDIT
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var book = await _db.Books.SingleOrDefaultAsync(m => m.Id == id);
            if (book == null)
            {
                return NotFound();
            }
            return View(book);
        }

        //EDIT VIEW-POST ACTION METHOD
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(Book book)
        {
            if (ModelState.IsValid)
            {
                //_db.Update(book); DISATAVGE it updates all coloums when maybe you dont want all
                var BookFromDb = await _db.Books.FirstOrDefaultAsync(b => b.Id == book.Id);
                BookFromDb.Name = book.Name;
                BookFromDb.Author = book.Author;
                BookFromDb.Price = book.Price;

                await _db.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View();
        }

        //GET ACTION METHOD FOR Delete
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var book = await _db.Books.SingleOrDefaultAsync(m => m.Id == id);
            if (book == null)
            {
                return NotFound();
            }
            return View(book);
        }

        //Delete VIEW-POST ACTION METHOD
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Remove(int? id)
        {
            var book = await _db.Books.SingleOrDefaultAsync(m => m.Id == id);
            _db.Books.Remove(book);

            await _db.SaveChangesAsync();
            return RedirectToAction(nameof(Index));

        }

        //GET ACTION METHOD FOR Details
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var book = await _db.Books.SingleOrDefaultAsync(m => m.Id == id);
            if (book == null)
            {
                return NotFound();
            }
            return View(book);
        }

     

    }
}
