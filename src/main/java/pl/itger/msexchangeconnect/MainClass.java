/*
 * Copyright (C) 2019 Piotr Zerynger
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package pl.itger.msexchangeconnect;

import java.io.IOException;
import java.net.URISyntaxException;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.joda.time.DateTime;

public class MainClass {

    private static Logger logger = Logger.getLogger(MainClass.class.getName());

    public static void main(String[] args) {
        OutlookReader exchangeCli;
        try {
            exchangeCli = new OutlookReader(DateTime.now());
        } catch (IOException | URISyntaxException ex) {
            logger.log(Level.SEVERE, null, ex);
            return;
        }
        ExecutorService executor = Executors.newFixedThreadPool(1);
        Future<?> future = executor.submit(exchangeCli);
        //List<String> smsy = new ArrayList<String>();
        logger.info("future executor.submitted ... ");
        try {
            future.get(60, TimeUnit.SECONDS);
        } catch (TimeoutException te) {
            logger.info("timeOut ... ");
            future.cancel(true);
        } catch (InterruptedException | ExecutionException ex) {
            logger.log(Level.SEVERE, null, ex);
            Thread.currentThread().interrupt();
        }
        exchangeCli.shutdownPool(executor);
    }
}
