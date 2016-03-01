package com.hotcocoacup.mobiletools.xlstoresouces;

import com.hotcocoacup.mobiletools.xlstoresouces.model.KeyValuePair;

import java.io.IOException;
import java.io.Writer;
import java.util.List;
import java.util.Map;

public interface Processor {

	void process (Writer outputStream, Map<String, List<KeyValuePair>> keyValuePair) throws IOException;
	
}
