import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import model.Event;
import org.junit.Test;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class JacksonTest {

    @Test

    public void TestJsonSerialize(){

        //SimpleDateFormat df = new SimpleDateFormat("yyyy/MM/DD hh:mm:ss");
        //String toParse = "2019/08/19 16:28:00";
        Date date = new Date();
        /*try {
            date = df.parse(toParse);
        } catch (ParseException e) {
            e.printStackTrace();
        }*/
        Event event = new Event();
        event.name = "party";
        event.eventDate = date;
        try {
            String result = new ObjectMapper().writeValueAsString(event);
            System.out.println(result);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }

    }

}
