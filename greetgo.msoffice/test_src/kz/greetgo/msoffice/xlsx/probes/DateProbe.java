package kz.greetgo.msoffice.xlsx.probes;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

public class DateProbe {
  private static final SimpleDateFormat F = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
  
  public static void main(String[] args) {
    {
      Calendar c = new GregorianCalendar();
      c.set(Calendar.YEAR, 1970);
      c.set(Calendar.MONTH, 1);
      c.set(Calendar.DAY_OF_MONTH, 1);
      
      System.out.println(F.format(c.getTime()) + " ||| " + c.getTimeInMillis());
      
      c.add(Calendar.DAY_OF_YEAR, 40868 - (70 * 365 + 50));
      
      System.out.println(F.format(c.getTime()) + " ||| " + c.getTimeInMillis());
    }
    {
      Calendar c = new GregorianCalendar();
      c.set(Calendar.YEAR, 1970);
      c.set(Calendar.MONTH, 1);
      c.set(Calendar.DAY_OF_MONTH, 1);
      
      c.add(Calendar.DAY_OF_YEAR, 42811 - (70 * 365 + 50));
      
      System.out.println(F.format(c.getTime()) + " ||| " + c.getTimeInMillis());
    }
    System.out.println(F.format(new Date(0L)));
    System.out.println("-------------------------------------------------");
    {
      Calendar c = new GregorianCalendar();
      c.set(Calendar.HOUR_OF_DAY, 0);
      c.set(Calendar.MINUTE, 0);
      c.set(Calendar.SECOND, 0);
      c.set(Calendar.MILLISECOND, 0);
      
      long t1 = c.getTimeInMillis();
      
      System.out.println(F.format(c.getTime()));
      
      c.set(Calendar.HOUR_OF_DAY, 24);
      c.set(Calendar.MINUTE, 0);
      c.set(Calendar.SECOND, 0);
      c.set(Calendar.MILLISECOND, 0);
      
      long t2 = c.getTimeInMillis();
      
      System.out.println(F.format(c.getTime()));
      
      System.out.println("t2 - t1 = " + (t2 - t1) + " Ж " + (24 * 60 * 60 * 1000));
    }
    System.out.println("-------------------------------------------------");
    {
      Calendar c = new GregorianCalendar();
      c.setTime(new Date(0));
      
      System.out.println(F.format(c.getTime()));
      
      c.set(Calendar.YEAR, 2011);
      c.set(Calendar.MONTH, 11);
      c.set(Calendar.DAY_OF_MONTH, 28);
      
      System.out.println(F.format(c.getTime()));
      
      c.set(Calendar.HOUR_OF_DAY, 0);
      
      System.out.println(F.format(c.getTime()));
    }
  }
}
