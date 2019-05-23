package excel.application;


public class Lessons {

    public String NameOfCourse, DayOfWeek, LocalTime, NumberOfHours, StartDate;
    public Lessons (String NameOfCourse, String DayOfWeek, String LocalTime, String NumberOfHours, String StartDate){

        this.NameOfCourse = NameOfCourse;
        this.DayOfWeek = DayOfWeek;
        this.LocalTime = LocalTime;
        this.NumberOfHours = NumberOfHours;
        this.StartDate = StartDate;
    }
}
