using System.Data;
using Microsoft.Data.SqlClient;

namespace Excel_Data_Importer_WARS
{
    internal class Pracownik
    {
        public string Imie = string.Empty;
        public string Nazwisko = string.Empty;
        public string Akronim = string.Empty;
        public int Get_PraId(SqlConnection connection, SqlTransaction transaction)
        {
            using SqlCommand command = new(DbManager.Get_PRI_PraId, connection, transaction);
            if (string.IsNullOrEmpty(Akronim))
            {
                command.Parameters.Add("@Akronim", SqlDbType.Int).Value = -1;
            }
            else
            {
                command.Parameters.Add("@Akronim", SqlDbType.Int).Value = int.Parse(Akronim);
            }
            command.Parameters.Add("@PracownikImieInsert", SqlDbType.NVarChar, 50).Value = Imie;
            command.Parameters.Add("@PracownikNazwiskoInsert", SqlDbType.NVarChar, 50).Value = Nazwisko;
            int Pracid = command.ExecuteScalar() as int? ?? 0;
            //Incydent z 27.02.2025
            //HashSet<int> invalidIds = new HashSet<int>
            //{
            //    7, 34, 41, 62, 66, 67, 74, 83, 85, 105, 107, 110, 122, 128, 129, 130, 146, 151, 186, 188, 190,
            //    195, 196, 205, 210, 230, 273, 279, 305, 312, 321, 324, 327, 358, 363, 372, 376, 387, 405, 408,
            //    410, 416, 419, 427, 430, 431, 434, 439, 444, 466, 477, 483, 492, 519, 556, 560, 575, 580, 589,
            //    590, 592, 597, 603, 618, 621, 630, 633, 639, 653, 655, 656, 659, 667, 672, 675, 683, 684, 696,
            //    699, 706, 708, 713, 714, 716, 722, 727, 736, 752, 757, 775, 778, 810, 829, 830, 831, 833, 837,
            //    856, 866, 867, 871, 877, 879, 889, 892, 902, 922, 932, 940, 955, 1000, 1009, 1052, 1061, 1065,
            //    1089, 1095, 1106, 1110, 1138, 1142, 1216, 1232, 1236, 1240, 1242, 1244, 1247, 1268, 1276, 1284,
            //    1301, 1307, 1320, 1327, 1330, 1344, 1352, 1365, 1409, 1432, 1438, 1456, 1457, 1461, 1489, 1521,
            //    1526, 1528, 1533, 1564, 1568, 1579, 1586, 1617, 1686, 1700, 1749, 1751, 1805, 1808, 1837, 1850,
            //    1851, 1852, 1878, 1879, 1882, 1886, 1891, 1892, 1938, 1951, 1967, 1981, 2005, 2006, 2011, 2039,
            //    2056, 2081, 2106, 2120, 2146, 2212, 2262, 2275, 2316, 2325, 2329, 2334, 2335, 2386, 2387, 2408,
            //    2421, 2440, 2448, 2517, 2546, 2593, 2612, 2683, 2759, 2799, 2800, 2801, 2832, 2837, 2842, 2891,
            //    2910, 2917, 2920, 2927, 2957, 2962, 2974, 3018, 3033, 3061, 3066, 3102, 3146, 3162, 3203, 3233,
            //    3242, 3249, 3256, 3274, 3287, 3330, 3362, 3364, 3375, 3400, 3427, 3448, 3454, 3476, 3510, 3545,
            //    3584, 3600, 3608, 3610, 3656, 3658
            //};
            //if (invalidIds.Contains(Pracid))
            //{
            //    throw new Exception("Temporary nie dotykać tego usera");
            //}
            return Pracid;
        }
    }
}