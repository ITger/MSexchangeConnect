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
import java.net.URI;
import java.net.URISyntaxException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;
import java.util.logging.Logger;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;
import microsoft.exchange.webservices.data.search.filter.SearchFilter.SearchFilterCollection;
import org.joda.time.DateTime;

public class OutlookReader implements Runnable {

    private static final Logger LOGGER = Logger.getLogger(OutlookReader.class.getName());

    public static void shutdownPool(final ExecutorService pool) {
        pool.shutdown();
        try {
            if (!pool.awaitTermination(10, TimeUnit.SECONDS)) {
                pool.shutdownNow();
                if (!pool.awaitTermination(10, TimeUnit.SECONDS)) {
                    LOGGER.info("shutdownNow: time out error");
                }
            }
        } catch (final InterruptedException e) {
            pool.shutdownNow();
            LOGGER.log(Level.SEVERE, e.getMessage(), e);
            Thread.currentThread().interrupt();
        }
    }

    private ExchangeService service;
    private DateTime fromDate;
    private final String folderName;

    /**
     *
     * @param fromDate
     * @throws IOException
     * @throws URISyntaxException
     */
    public OutlookReader(final DateTime fromDate) throws IOException, URISyntaxException {
        this.fromDate = fromDate;
        this.folderName = "subFolderForTests";
        this.service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        connect();
    }

    /**
     * Connects to MS Exchange Service
     *
     * @throws URISyntaxException
     * @throws IOException
     */
    private void connect() throws URISyntaxException, IOException {
        LOGGER.info("connecting to ExchangeService");
        // the NewClass class contains hardcoded login and password
        // data as public static strings.
        this.service.setCredentials(new WebCredentials(
                NewClass.login,
                NewClass.passwd));
        this.service.setUrl(new URI("https://outlook.live.com/ews/exchange.asmx"));
        this.service.setTimeout(Integer.parseInt("60000"));
    }

    public void run() {
        final ItemView view = new ItemView(100);
        Folder subFolder = null;
        try {
            Folder rootFolder = Folder.bind(this.service, WellKnownFolderName.Inbox);
            for (final Folder folder : rootFolder.findFolders(new FolderView(500))) {
                if (folder.getDisplayName().contains(this.folderName)) {
                    subFolder = Folder.bind(service, folder.getId());
                    break;
                }
            }
        } catch (final ServiceLocalException e1) {
            LOGGER.log(Level.SEVERE, e1.getMessage(), e1);
            return;
        } catch (final Exception e1) {
            LOGGER.log(Level.SEVERE, e1.getMessage(), e1);
            return;
        }
        if (subFolder == null) {
            LOGGER.log(Level.SEVERE, "Folder ".concat(this.folderName).concat(" does not exist."));
            return;
        }
        try {
            view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
        } catch (final ServiceLocalException e1) {
            LOGGER.log(Level.SEVERE, e1.getMessage(), e1);
            return;
        }
        view.setPropertySet(new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived));

        SearchFilterCollection sfc = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.ContainsSubstring(ItemSchema.Subject, "testing ews"),
                //new SearchFilter.ContainsSubstring(ItemSchema.Body, ""),
                new SearchFilter.IsGreaterThan(EmailMessageSchema.DateTimeCreated, fromDate.toDate()));

        while (!Thread.interrupted()) {
            try {
                final FindItemsResults<Item> findResults = service.findItems(
                        subFolder.getId(), sfc, view);
                LOGGER.log(Level.INFO, "total ".concat(Integer.toString(findResults.getTotalCount())));
                if (Thread.interrupted()) {
                    LOGGER.log(Level.INFO, "Thread.interrupted");
                    break;
                }
                if (findResults.getTotalCount() > 0) {
                    service.loadPropertiesForItems(findResults, PropertySet.FirstClassProperties);
                    for (final Item item : findResults) {
                        try {
                            LOGGER.log(Level.INFO, item.getBody().toString());
                        } catch (final ServiceLocalException e) {
                            LOGGER.log(Level.SEVERE, e.getLocalizedMessage(), e);
                        }
                    }
                    break;
                }
            } catch (final Exception e) {
                LOGGER.log(Level.SEVERE, e.getLocalizedMessage(), e);
            }
            try {
                TimeUnit.SECONDS.sleep(1);
            } catch (final InterruptedException e) {
                LOGGER.info("InterruptedException ");
                Thread.currentThread().interrupt();
                break;
            }
        }
        this.service = null;
        return;
    }
}
