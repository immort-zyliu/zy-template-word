package pers.lzy.template.word.test.pojo;

public class OutAddress {


        private String name;
        private Integer code;

        public OutAddress(String name, Integer code) {
            this.name = name;
            this.code = code;
        }

        public OutAddress() {
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Integer getCode() {
            return code;
        }

        public void setCode(Integer code) {
            this.code = code;
        }

        @Override
        public String toString() {
            return "Out{" +
                    "name='" + name + '\'' +
                    ", code=" + code +
                    '}';
        }
    }

