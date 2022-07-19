namespace MealVouchers
{
    public static class Extensions
    {
        public static int IndexOf<T>(this IReadOnlyList<T> objects, T elementToFind)
        {
            int i = 0;
            foreach (T element in objects)
            {
                if (Equals(element, elementToFind))
                    return i;
                i++;
            }
            return -1;
        }
    }
}
