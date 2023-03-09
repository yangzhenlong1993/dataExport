package utils;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
public @interface Excel {
    KeyGeneratorType keyType() default KeyGeneratorType.AUTO;

    int order() default 0;

    String columnName() default "";

    String fileName() default "excel";
}
