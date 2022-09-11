namespace EasyUseExcel.Attribute
{
    public class DisplayAttribute : System.Attribute
    {
        public DisplayAttribute(string Name)
        {
            this.Name = Name;
        }

        public virtual string Name { get; set; }
    }
}
