namespace PECTutorial.Models
{
    public class TestModel
    {
    }

    public class ItemViewModel
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public bool IsSelected { get; set; }
    }

    public class ListBindingViewModel
    {
        public IList<ItemViewModel> Items { get; set; } = new List<ItemViewModel>();
    }

    public class PersonAddress
    {
        public string City { get; set; }
        public string Country { get; set; }
    }
}
