package model;

import lombok.Builder;

@Builder
public class User {
    private final Integer code = 200;
    private String username;
    private String password;
}

