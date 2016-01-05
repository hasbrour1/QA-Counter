using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
 * Analyte Class
 * Object class for each analyte
 * records each count for each month
 *
 */ 

namespace QA_Counter
{
    public class Analyte
    {
        public String name;
        public String matrix;
        public String detMatrix;
        public int cJan;
        public int cFeb;
        public int cMar;
        public int cApr;
        public int cMay;
        public int cJun;
        public int cJul;
        public int cAug;
        public int cSep;
        public int cOct;
        public int cNov;
        public int cDec;

        public Analyte(String na, String m, String d, int c, int month)
        {
            name = na;
            matrix = m;
            detMatrix = d;
            cJan = 0;
            cFeb = 0;
            cMar = 0;
            cApr = 0;
            cMay = 0;
            cJun = 0;
            cJul = 0;
            cAug = 0;
            cSep = 0;
            cOct = 0;
            cNov = 0;
            cDec = 0;
            addCount(c, month);
        }

        public void addCount(int c, int month)
        {

            switch (month)
            {
                case 0:
                    cJan += c;
                    break;
                case 1:
                    cFeb += c;
                    break;
                case 2:
                    cMar += c;
                    break;
                case 3:
                    cApr += c;
                    break;
                case 4:
                    cMay += c;
                    break;
                case 5:
                    cJun += c;
                    break;
                case 6:
                    cJul += c;
                    break;
                case 7:
                    cAug += c;
                    break;
                case 8:
                    cSep += c;
                    break;
                case 9:
                    cOct += c;
                    break;
                case 10:
                    cNov += c;
                    break;
                case 11:
                    cDec += c;
                    break;
                default:
                    break;
            }           
        }

        public int getCount(int month)
        {
            switch (month)
            {
                case 0:
                    return cJan;
                case 1:
                    return cFeb;
                case 2:
                    return cMar;
                case 3:
                    return cApr;
                case 4:
                    return cMay;
                case 5:
                    return cJun;
                case 6:
                    return cJul;
                case 7:
                    return cAug;
                case 8:
                    return cSep;
                case 9:
                    return cOct;
                case 10:
                    return cNov;
                case 11:
                    return cDec;
                default:
                    return 0;
            }         
        }

        public string getMatrix()
        {
            return matrix;
        }

        public string getDetMatrix()
        {
            return detMatrix;
        }

        public void setDetMatrix(String str)
        {
            detMatrix = str;
        }

        public string getName()
        {
            return name;
        }
    }
}
