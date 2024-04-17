package entity;

import lombok.Data;
import lombok.experimental.Accessors;

@Accessors(chain = true)
@Data
public class Normal_Student {
    String building;
    String room_number;
    String bed_number;
    String score;
}
