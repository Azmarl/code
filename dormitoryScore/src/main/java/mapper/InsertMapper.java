package mapper;

import entity.Normal_Student;
import org.apache.ibatis.annotations.Insert;

public interface InsertMapper {
    @Insert("insert into seventeen_pf  (building, room_number, bed_number, score) values (#{building}, #{room_number}, #{bed_number}, #{score})")
    int addNormalStudent(Normal_Student student);
}
