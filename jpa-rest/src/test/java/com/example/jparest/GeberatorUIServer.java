package com.example.jparest;

import com.github.davidfantasy.mybatisplus.generatorui.GeneratorConfig;
import com.github.davidfantasy.mybatisplus.generatorui.MybatisPlusToolsApplication;
import com.github.davidfantasy.mybatisplus.generatorui.mbp.NameConverter;

public class GeberatorUIServer {

    public static void main(String[] args) {
        GeneratorConfig config = GeneratorConfig.builder().jdbcUrl("jdbc:postgresql://psql-sdm-dev-southeastasia-001.postgres.database.azure.com:6432/psc?useUnicode=true&characterEncoding=utf-8&prepareThreshold=0")
                .userName("pscd")
                .password("x0DIlC8oij5s")
                .driverClassName("org.postgresql.Driver")
                //数据库schema，MSSQL,PGSQL,ORACLE,DB2类型的数据库需要指定
                .schemaName("sdb")
                //数据库表前缀，生成entity名称时会去掉(v2.0.3新增)
//                .tablePrefix("t_")
                //如果需要修改entity及其属性的命名规则，以及自定义各类生成文件的命名规则，可自定义一个NameConverter实例，覆盖相应的名称转换方法，详细可查看该接口的说明：                
                .nameConverter(new NameConverter() {
                    /**
                     * 自定义Service类文件的名称规则，entityName是NameConverter.entityNameConvert处理表名后的返回结果，如有特别的需求可以自定义实现
                     */
                    @Override
                    public String serviceNameConvert(String entityName) {
                        return entityName + "Service";
                    }

                    /**
                     * 自定义Controller类文件的名称规则
                     */
                    @Override
                    public String controllerNameConvert(String entityName) {
                        return entityName + "Action";
                    }
                })
                //所有生成的java文件的父包名，后续也可单独在界面上设置
                .basePackage("com.github.davidfantasy.mybatisplustools.example")
                .port(8068)
                .build();
        MybatisPlusToolsApplication.run(config);
    }

}