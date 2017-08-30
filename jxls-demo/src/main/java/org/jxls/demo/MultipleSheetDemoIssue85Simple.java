package org.jxls.demo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.jxls.common.Context;
import org.jxls.demo.model.Department;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;

/**
 * https://bitbucket.org/leonate/jxls/issues/85/wrong-sum-in-multisheet
 * 
 * @author Marcus W.
 */
public class MultipleSheetDemoIssue85Simple {
    private static String template = "issue85-simple.xlsx";
    private static String output = "target/issue85-simple_output.xlsx";

    public static void main(String[] args) throws IOException {
    	EachIfCommandDemo.make300 = true;
        List<Department> departments = EachIfCommandDemo.createDepartments();
        try (InputStream is = MultipleSheetDemoIssue85Simple.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Context context = PoiTransformer.createInitialContext();
                context.putVar("departments", departments.get(1).getStaff());
                JxlsHelper.getInstance().setUseFastFormulaProcessor(false).processTemplate(is, os, context);
            }
        }
        System.out.println("finished");
        // see wrong sum in 2nd sheet cell F13 (also wrong on other sheets)
    }
}
