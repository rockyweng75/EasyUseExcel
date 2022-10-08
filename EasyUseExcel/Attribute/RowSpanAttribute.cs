namespace EasyUseExcel.Attribute
{
    public class OrderAttribute : System.Attribute
    {
        public OrderAttribute(int index)
        {
            Index = index;
        }

        public int Index { get; set; }
    }
}
