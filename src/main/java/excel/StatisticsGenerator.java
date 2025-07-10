package excel;

import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

public class StatisticsGenerator {

    public static void generateStatistics(List<Entry> entries) {
        System.out.println("\n📊 Ranking employees by total hours:");
        generateEmployeeRanking(entries);

        System.out.println("\n📆 Ranking months by total hours:");
        generateMonthRanking(entries);

        System.out.println("\n🔥 Top 10 most productive days:");
        generateTopDays(entries);
    }

    private static void generateEmployeeRanking(List<Entry> entries) {
        Map<String, Double> hoursByEmployee = new HashMap<>();

        for (Entry e : entries) {
            hoursByEmployee.merge(e.employeeName(), e.hours(), Double::sum);
        }

        hoursByEmployee.entrySet().stream()
                .sorted(Map.Entry.<String, Double>comparingByValue().reversed())
                .forEach(entry ->
                        System.out.printf("%s - %.2f hours%n", entry.getKey(), entry.getValue()));
    }

    private static void generateMonthRanking(List<Entry> entries) {
        Map<String, Double> hoursByMonth = new HashMap<>();
        SimpleDateFormat monthYearFormat = new SimpleDateFormat("MMMM yyyy", Locale.ENGLISH);

        for (Entry e : entries) {
            String month = monthYearFormat.format(e.date());
            hoursByMonth.merge(month, e.hours(), Double::sum);
        }

        List<Map.Entry<String, Double>> sorted = hoursByMonth.entrySet().stream()
                .sorted(Map.Entry.<String, Double>comparingByValue().reversed())
                .toList();

        int i = 1;
        for (Map.Entry<String, Double> entry : sorted) {
            System.out.printf("%d. %s - %.0f hours%n", i++, entry.getKey(), entry.getValue());
        }
    }

    private static void generateTopDays(List<Entry> entries) {
        Map<String, Double> hoursByDay = new HashMap<>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy", Locale.ENGLISH);

        for (Entry e : entries) {
            String day = dateFormat.format(e.date());
            hoursByDay.merge(day, e.hours(), Double::sum);
        }

        List<Map.Entry<String, Double>> top10 = hoursByDay.entrySet().stream()
                .sorted(Map.Entry.<String, Double>comparingByValue().reversed())
                .limit(10)
                .toList();

        int i = 1;
        for (Map.Entry<String, Double> entry : top10) {
            System.out.printf("%d. %s - %.0f hours%n", i++, entry.getKey(), entry.getValue());
        }
    }
}
