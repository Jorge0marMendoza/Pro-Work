﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BookListRazor.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace BookListRazor.Pages.BookList
{
    public class EditModel : PageModel
    {
        private ApplicationDbContext _db;
        public EditModel(ApplicationDbContext db)
        {
            _db = db;
            
        }

        [TempData]
        public string Message { get; set; }

        [BindProperty]
        public Book Book { get; set; }
        public async Task OnGet(int id)
        {
            Book = await _db.Book.FindAsync(id);

        }

        public async Task<IActionResult> OnPost()
        {
            if (ModelState.IsValid)
            {
                var BookFromBD = await _db.Book.FindAsync(Book.Id);
                BookFromBD.Name = Book.Name;
                BookFromBD.Author = Book.Author;
                BookFromBD.ISBN = Book.ISBN;

                await _db.SaveChangesAsync();
                Message = "Book has been updated successfully";
                return RedirectToPage("Index");
            }

            return RedirectToPage();
        }
    }
}