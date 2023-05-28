package org.example.Main;

import org.example.Parsing.Parsing;
import org.example.SendEmail.SendEmail;

import java.util.Calendar;
import java.util.Date;
import java.util.Timer;
import java.util.TimerTask;

public class Main {
    public static void main(String[] args) {
        Timer timer = new Timer();

        Calendar currentDate = Calendar.getInstance();
        Date now = currentDate.getTime();

        Calendar scheduledTime = Calendar.getInstance();
        scheduledTime.set(Calendar.HOUR_OF_DAY, 8);
        scheduledTime.set(Calendar.MINUTE, 0);
        scheduledTime.set(Calendar.SECOND, 0);

        if (now.compareTo(scheduledTime.getTime()) >= 0) {
            scheduledTime.add(Calendar.DATE, 1);
        }

        long initialDelay = scheduledTime.getTimeInMillis() - now.getTime();

        timer.scheduleAtFixedRate(new CustomTask(), initialDelay, 1000 * 3600 * 24);
    }

    static class CustomTask extends TimerTask {
        @Override
        public void run() {
            Parsing.parse();
            System.out.println("Сообщение с Excel-файлом будет отправлено через 30 секунд");
            try {
                Thread.sleep(1000 * 30);
                SendEmail.send();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
