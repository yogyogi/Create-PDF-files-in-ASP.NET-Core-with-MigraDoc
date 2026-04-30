using Microsoft.AspNetCore.Mvc;
using PECTutorial.Models;

namespace PECTutorial.Controllers
{
    public class TestController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult EditList()
        {
            var model = new ListBindingViewModel
            {
                Items = new List<ItemViewModel>
            {
                new ItemViewModel { Id = 1, Name = "Item A" },
                new ItemViewModel { Id = 2, Name = "Item B" },
                new ItemViewModel { Id = 3, Name = "Item C" }
            }
            };
            return View(model);
        }

        [HttpPost]
        public IActionResult EditList(ListBindingViewModel model)
        {
            if (ModelState.IsValid)
            {
                foreach (var item in model.Items)
                {
                    System.Diagnostics.Debug.WriteLine($"Item ID: {item.Id}, Name: {item.Name}, Selected: {item.IsSelected}");
                }
                return RedirectToAction("Success");
            }
            return View(model);
        }

        public IActionResult Address()
        {
            List<PersonAddress> addresses = new List<PersonAddress>()
            {
                new PersonAddress { City = "New York", Country = "USA" },
                new PersonAddress { City = "London", Country = "UK" },
                new PersonAddress { City = "Tokyo", Country = "Japan" } 
            };
            return View();
        }

        [HttpPost]
        public IActionResult Address(List<PersonAddress> address) => View(address);
    }
}
